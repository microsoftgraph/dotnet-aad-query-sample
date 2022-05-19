// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using Microsoft.Graph;
using Microsoft.Toolkit.Mvvm.ComponentModel;
using Microsoft.Toolkit.Mvvm.Input;
using MsGraph_Samples.Services;
using System.Diagnostics;
using System.Net;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace MsGraph_Samples.ViewModels
{
    [ObservableObject]
    public partial class MainViewModel
    {
        private readonly IAuthService _authService;
        private readonly IGraphDataService _graphDataService;

        private readonly Stopwatch _stopWatch = new();
        public long ElapsedMs => _stopWatch.ElapsedMilliseconds;

        [ObservableProperty]
        private bool _isBusy;

        [ObservableProperty]
        private string? _userName;

        public string? LastUrl => _graphDataService.LastUrl;

        public static IReadOnlyList<string> Entities => new[] { "Users", "Groups", "Applications", "Devices" };

        [ObservableProperty]
        private string _selectedEntity = "Users";

        [ObservableProperty]
        private DirectoryObject? _selectedObject;

        [ObservableProperty]
        private IEnumerable<DirectoryObject>? _directoryObjects;

        #region OData Operators

        public string[] SplittedSelect => Select.Split(',', StringSplitOptions.TrimEntries);

        [ObservableProperty]
        public string _select = "id,displayName,mail,userPrincipalName";

        [ObservableProperty]
        public string _filter = string.Empty;

        [ObservableProperty]
        public string _orderBy = string.Empty;

        private string _search = string.Empty;
        public string Search
        {
            get => _search;
            set
            {
                if (_search == value)
                    return;

                _search = FixSearchSyntax(value);
                OnPropertyChanged();
            }
        }

        private static string FixSearchSyntax(string searchValue)
        {
            if (searchValue.Contains('"'))
                return searchValue; // Assume already correctly formatted

            var elements = searchValue.Trim().Split(' ');
            var sb = new StringBuilder(elements.Length);

            foreach (var element in elements)
            {
                string? newElement;

                if (element.Contains(':'))
                    newElement = $"\"{element}\""; // Search clause needs to be wrapped by double quotes
                else if (element.In("AND", "OR"))
                    newElement = $" {element.ToUpperInvariant()} "; // [AND, OR] Operators need to be uppercase
                else
                    newElement = element;

                sb.Append(newElement);
            }

            return sb.ToString();
        }

        #endregion

        public MainViewModel(IAuthService authService, IGraphDataService graphDataService)
        {
            _authService = authService;
            _graphDataService = graphDataService;
            Init().Await();
        }

        public async Task Init()
        {
            var user = await _graphDataService.GetMe("displayName");
            UserName = user.DisplayName;
            await Load();
        }
        

        [ICommand]
        //TODO implement [AlsoNotifyCanExecuteFor(nameof(GraphExplorerCommand))] when upgrading to MMVM Toolkit 8.0
        private async Task Load()
        {
            await IsBusyWrapper(async () =>
            {
                DirectoryObjects = SelectedEntity switch
                {
                    "Users" => await _graphDataService.GetUsersAsync(Select, Filter, OrderBy, Search),
                    "Groups" => await _graphDataService.GetGroupsAsync(Select, Filter, OrderBy, Search),
                    "Applications" => await _graphDataService.GetApplicationsAsync(Select, Filter, OrderBy, Search),
                    "Devices" => await _graphDataService.GetDevicesAsync(Select, Filter, OrderBy, Search),
                    _ => throw new NotImplementedException("Can't find selected entity")
                };
            });
        }

        [ICommand]
        //TODO: implement [ICommand(CanExecute = SelectedObject is not null)] after upgrade to MVVM Toolkit 8.0
        private async Task DrillDown()
        {
            if (SelectedObject is null)
                return;

            await IsBusyWrapper(async () =>
            {
                DirectoryObjects = SelectedEntity switch
                {
                    "Users" => await _graphDataService.GetTransitiveMemberOfAsGroupsAsync(SelectedObject.Id),
                    "Groups" => await _graphDataService.GetTransitiveMembersAsUsersAsync(SelectedObject.Id),
                    "Applications" => await _graphDataService.GetAppOwnersAsUsersAsync(SelectedObject.Id),
                    "Devices" => await _graphDataService.GetTransitiveMemberOfAsGroupsAsync(SelectedObject.Id),
                    _ => null
                };

                SelectedEntity = DirectoryObjects switch
                {
                    IGraphServiceUsersCollectionPage => "Users",
                    IGraphServiceGroupsCollectionPage => "Groups",
                    IGraphServiceApplicationsCollectionPage => "Applications",
                    IGraphServiceDevicesCollectionPage => "Devices",
                    _ => SelectedEntity,
                };
            });
        }

        [ICommand]
        private Task Sort(DataGridSortingEventArgs? e)
        {
            ArgumentNullException.ThrowIfNull(e);

            OrderBy = $"{e.Column.Header}";
            e.Handled = true;
            return Load();
        }

        [ICommand] //TODO: implement [ICommand(CanExecute = LastUrl is not null)] after upgrade to MVVM Toolkit 8.0
        private void GraphExplorer()
        {
            if (LastUrl is null) return;

            var geBaseUrl = "https://developer.microsoft.com/graph/graph-explorer";
            var graphUrl = "https://graph.microsoft.com";
            var version = "v1.0";
            var startOfQuery = LastUrl.NthIndexOf('/', 4) + 1;
            var encodedUrl = WebUtility.UrlEncode(LastUrl[startOfQuery..]);
            var encodedHeaders = "W3sibmFtZSI6IkNvbnNpc3RlbmN5TGV2ZWwiLCJ2YWx1ZSI6ImV2ZW50dWFsIn1d"; // ConsistencyLevel = eventual

            var url = $"{geBaseUrl}?request={encodedUrl}&method=GET&version={version}&GraphUrl={graphUrl}&headers={encodedHeaders}";

            var psi = new ProcessStartInfo { FileName = url, UseShellExecute = true };
            System.Diagnostics.Process.Start(psi);
        }

        [ICommand]
        private void Logout()
        {
            _authService.Logout();
            App.Current.Shutdown();
        }

        private async Task IsBusyWrapper(Func<Task> task)
        {
            IsBusy = true;
            _stopWatch.Restart();

            try
            {
                await task();
            }
            catch (ServiceException ex)
            {
                MessageBox.Show(ex.Message, ex.Error.Message);
            }
            finally
            {
                _stopWatch.Stop();
                OnPropertyChanged(nameof(ElapsedMs));
                OnPropertyChanged(nameof(LastUrl));
                IsBusy = false;
            }
        }
    }
}