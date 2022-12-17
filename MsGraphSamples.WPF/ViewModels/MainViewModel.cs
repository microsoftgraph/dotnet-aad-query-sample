// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Microsoft.Graph.Models;
using Microsoft.Graph.Models.ODataErrors;
using MsGraph_Samples.Services;
using System.Diagnostics;
using System.Net;
using System.Text;
using System.Windows;
using System.Windows.Controls;

namespace MsGraph_Samples.ViewModels
{
    public partial class MainViewModel : ObservableObject
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
        [NotifyCanExecuteChangedFor(nameof(LaunchGraphExplorerCommand))]
        private BaseCollectionPaginationCountResponse? _directoryObjects;

        #region OData Operators

        public string[] SplittedSelect => Select.Split(',', StringSplitOptions.TrimEntries);

        [ObservableProperty]
        public string _select = "id,displayName,mail,userPrincipalName";

        [ObservableProperty]
        public string _filter = string.Empty;

        public string[] SplittedOrderBy => OrderBy.Split(',', StringSplitOptions.TrimEntries);

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
            var user = await _graphDataService.GetMe(new[] { "displayName" });
            UserName = user.DisplayName;
            await Load();
        }


        [RelayCommand]
        private async Task Load()
        {
            await IsBusyWrapper(async () =>
            {
                DirectoryObjects = SelectedEntity switch
                {
                    "Users" => await _graphDataService.GetUsersAsync(SplittedSelect, Filter, SplittedOrderBy, Search),
                    "Groups" => await _graphDataService.GetGroupsAsync(SplittedSelect, Filter, SplittedOrderBy, Search),
                    "Applications" => await _graphDataService.GetApplicationsAsync(SplittedSelect, Filter, SplittedOrderBy, Search),
                    "Devices" => await _graphDataService.GetDevicesAsync(SplittedSelect, Filter, SplittedOrderBy, Search),
                    _ => throw new NotImplementedException("Can't find selected entity")
                };
            });
        }

        private bool CanDrillDown() => SelectedObject is not null;
        [RelayCommand(CanExecute = nameof(CanDrillDown))]
        private async Task DrillDown()
        {
            ArgumentNullException.ThrowIfNull(SelectedObject);

            await IsBusyWrapper(async () =>
            {
                OrderBy = string.Empty;
                Filter = string.Empty;
                Search = string.Empty;

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
                    UserCollectionResponse => "Users",
                    GroupCollectionResponse => "Groups",
                    ApplicationCollectionResponse => "Applications",
                    DeviceCollectionResponse => "Devices",
                    _ => SelectedEntity,
                };
            });
        }

        [RelayCommand]
        private Task Sort(DataGridSortingEventArgs? e)
        {
            ArgumentNullException.ThrowIfNull(e);

            OrderBy = $"{e.Column.Header}";
            e.Handled = true;
            return Load();
        }

        private bool CanLaunchGraphExplorer() => LastUrl is not null;
        [RelayCommand(CanExecute = nameof(CanLaunchGraphExplorer))]
        private void LaunchGraphExplorer()
        {
            ArgumentNullException.ThrowIfNull(LastUrl);

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

        [RelayCommand]
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
            catch (ODataError ex)
            {
                Task.Run(() => MessageBox.Show(ex.Message, ex.Error.Message)).Await();
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