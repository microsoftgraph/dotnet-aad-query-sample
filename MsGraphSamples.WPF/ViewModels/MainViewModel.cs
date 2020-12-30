// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Graph;
using MsGraph_Samples.MVVM;
using MsGraph_Samples.Services;

namespace MsGraph_Samples.ViewModels
{
    public class MainViewModel : ObservableObject
    {
        private readonly IAuthService _authService;
        private readonly IGraphDataService _graphDataService;

        private readonly Stopwatch _stopWatch = new();

        private bool _isBusy;
        public bool IsBusy
        {
            get => _isBusy;
            set => Set(ref _isBusy, value);
        }

        private string? _userName;
        public string? UserName
        {
            get => _userName;
            set => Set(ref _userName, value);
        }

        public string? LastUrl => _graphDataService.LastUrl;

        public static IReadOnlyList<string> Entities => new[] { "Users", "Groups", "Applications", "Devices" };
        private string _selectedEntity = "Users";
        public string SelectedEntity
        {
            get => _selectedEntity;
            set => Set(ref _selectedEntity, value);
        }

        private DirectoryObject? _selectedObject = null;
        public DirectoryObject? SelectedObject
        {
            get => _selectedObject;
            set => Set(ref _selectedObject, value);
        }

        public long ElapsedMs => _stopWatch.ElapsedMilliseconds;

        private IEnumerable<DirectoryObject>? _directoryObjects;
        public IEnumerable<DirectoryObject>? DirectoryObjects
        {
            get => _directoryObjects;
            set => Set(ref _directoryObjects, value);
        }

        #region OData Operators

        private const string SelectDefaultValue = "id, displayName, mail, userPrincipalName";
        public string[] SplittedSelect { get; private set; } = SelectDefaultValue.Split(',', StringSplitOptions.TrimEntries);

        private string _select = SelectDefaultValue;
        public string Select
        {
            get => _select;
            set
            {
                if (Set(ref _select, value))
                    SplittedSelect = Select.Split(',', StringSplitOptions.TrimEntries);
            }
        }

        private string _filter = string.Empty;
        public string Filter
        {
            get => _filter;
            set => Set(ref _filter, value);
        }

        private string _orderBy = string.Empty;
        public string OrderBy
        {
            get => _orderBy;
            set => Set(ref _orderBy, value);
        }

        private string _search = string.Empty;
        public string Search
        {
            get => _search;
            set
            {
                if (_search == value)
                    return;

                _search = FixSearchSyntax(value);
                RaisePropertyChanged();
            }
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
            await LoadAction();

            var user = await _graphDataService.GetMe();
            UserName = user.DisplayName;
        }

        public AsyncRelayCommand LoadCommand => new(LoadAction);
        private async Task LoadAction()
        {
            IsBusy = true;
            _stopWatch.Restart();

            try
            {
                DirectoryObjects = SelectedEntity switch
                {
                    "Users" => await _graphDataService.GetUsersAsync(Select, Filter, OrderBy, Search),
                    "Groups" => await _graphDataService.GetGroupsAsync(Select, Filter, OrderBy, Search),
                    "Applications" => await _graphDataService.GetApplicationsAsync(Select, Filter, OrderBy, Search),
                    "Devices" => await _graphDataService.GetDevicesAsync(Select, Filter, OrderBy, Search),
                    _ => throw new NotImplementedException("Can't find selected entity")
                };
            }
            catch (ServiceException ex)
            {
                MessageBox.Show(ex.Message, ex.Error.Message);
            }
            finally
            {
                _stopWatch.Stop();
                RaisePropertyChanged(nameof(ElapsedMs));
                RaisePropertyChanged(nameof(LastUrl));
                RelayCommand.RaiseCanExecuteChanged();
                IsBusy = false;
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

        public AsyncRelayCommand DrillDownCommand => new(DrillDownAction);
        private async Task DrillDownAction()
        {
            if (SelectedObject == null) return;

            IsBusy = true;
            _stopWatch.Restart();

            Filter = string.Empty;
            Search = string.Empty;

            try
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
            }
            catch (ServiceException ex)
            {
                MessageBox.Show(ex.Message, ex.Error.Message);
            }
            finally
            {
                _stopWatch.Stop();
                RaisePropertyChanged(nameof(ElapsedMs));
                RaisePropertyChanged(nameof(LastUrl));
                AsyncRelayCommand.RaiseCanExecuteChanged();
                IsBusy = false;
            }
        }

        public AsyncRelayCommand<DataGridSortingEventArgs> SortCommand => new(SortAction);
        private Task SortAction(DataGridSortingEventArgs e)
        {
            OrderBy = $"{e.Column.Header}";
            e.Handled = true;
            return LoadAction();
        }

        public RelayCommand GraphExplorerCommand => new(GraphExplorerAction, () => LastUrl is not null);
        private void GraphExplorerAction()
        {
            if (LastUrl == null) return;

            var geBaseUrl = "https://developer.microsoft.com/en-us/graph/graph-explorer";
            var graphUrl = "https://graph.microsoft.com";
            var version = "v1.0";
            var encodedUrl = WebUtility.UrlEncode(LastUrl[(LastUrl.NthIndexOf('/', 4) + 1)..]);
            var encodedHeaders = "W3sibmFtZSI6IkNvbnNpc3RlbmN5TGV2ZWwiLCJ2YWx1ZSI6ImV2ZW50dWFsIn1d"; // ConsistencyLevel = eventual

            var url = $"{geBaseUrl}?request={encodedUrl}&method=GET&version={version}&GraphUrl={graphUrl}&headers={encodedHeaders}";

            var psi = new ProcessStartInfo { FileName = url, UseShellExecute = true };
            System.Diagnostics.Process.Start(psi);
        }

        public AsyncRelayCommand LogoutCommand => new(LogoutAction, () => UserName is not null);

        private async Task LogoutAction()
        {
            await _authService.Logout();
            App.Current.Shutdown();
        }
    }
}