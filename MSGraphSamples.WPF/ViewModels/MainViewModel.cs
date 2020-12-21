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

        private readonly Stopwatch _stopWatch = new Stopwatch();

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
            set
            {
                Set(ref _directoryObjects, value);
                SelectedEntity = DirectoryObjects switch
                {
                    GraphServiceUsersCollectionPage _ => "Users",
                    GraphServiceGroupsCollectionPage _ => "Groups",
                    GraphServiceApplicationsCollectionPage _ => "Applications",
                    GraphServiceDevicesCollectionPage _ => "Devices",
                    _ => SelectedEntity,
                };
            }
        }

        private string _select = "id, displayName, mail, userPrincipalName";
        public string Select
        {
            get => _select;
            set => Set(ref _select, value);
        }

        private string _filter = string.Empty;
        public string Filter
        {
            get => _filter;
            set => Set(ref _filter, value);
        }

        private string _search = string.Empty;
        public string Search
        {
            get => _search;
            set => Set(ref _search, value);
        }

        private string _orderBy = string.Empty;
        public string OrderBy
        {
            get => _orderBy;
            set => Set(ref _orderBy, value);
        }

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

        private AsyncRelayCommand? _loadCommand;
        public AsyncRelayCommand LoadCommand => _loadCommand ??= new AsyncRelayCommand(LoadAction);
        private async Task LoadAction()
        {
            FixSearchSyntax();

            IsBusy = true;
            _stopWatch.Restart();

            try
            {
                DirectoryObjects = SelectedEntity switch
                {
                    "Users" => await _graphDataService.GetUsersAsync(Filter, Search, Select, OrderBy),
                    "Groups" => await _graphDataService.GetGroupsAsync(Filter, Search, Select, OrderBy),
                    "Applications" => await _graphDataService.GetApplicationsAsync(Filter, Search, Select, OrderBy),
                    "Devices" => await _graphDataService.GetDevicesAsync(Filter, Search, Select, OrderBy),
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

        private void FixSearchSyntax()
        {
            if (Search.IsNullOrEmpty())
                return;

            if (Search.Contains('"'))
                return;

            var elements = Search.Split(' ');
            var sb = new StringBuilder(elements.Length);

            foreach (var element in elements)
            {
                var newElement = element.Contains(':') ?
                    $"\"{element}\"" : // Search clause needs to be wrapped by double quotes
                    $" {element.ToUpperInvariant()} "; // [AND, OR] operators need to be uppercase
                sb.Append(newElement);
            }

            Search = sb.ToString();
        }

        private AsyncRelayCommand? _drillDownCommand;
        public AsyncRelayCommand DrillDownCommand => _drillDownCommand ??= new AsyncRelayCommand(DrillDownAction);
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

        public RelayCommand<DataGridAutoGeneratingColumnEventArgs> AutoGeneratingColumn =>
            new RelayCommand<DataGridAutoGeneratingColumnEventArgs>(AutoGeneratingColumnAction);
        private void AutoGeneratingColumnAction(DataGridAutoGeneratingColumnEventArgs e)
        {
            if (!Select.IsNullOrEmpty())
            {
                e.Cancel = !e.PropertyName.In(Select.Split(','));
                e.Column.Width = new DataGridLength(1, DataGridLengthUnitType.Star);
            }
        }

        private AsyncRelayCommand<DataGridSortingEventArgs>? _sortCommand;
        public AsyncRelayCommand<DataGridSortingEventArgs> SortCommand => _sortCommand ??= new AsyncRelayCommand<DataGridSortingEventArgs>(SortAction);
        private Task SortAction(DataGridSortingEventArgs e)
        {
            OrderBy = $"{e.Column.Header}";
            e.Handled = true;
            return LoadAction();
        }

        private RelayCommand? _graphExplorerCommand;
        public RelayCommand GraphExplorerCommand => _graphExplorerCommand ??= new RelayCommand(GraphExplorerAction, () => LastUrl is not null);
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

        private AsyncRelayCommand? _logoutCommand;
        public AsyncRelayCommand LogoutCommand => _logoutCommand ??= new AsyncRelayCommand(LogoutAction, () => UserName is not null);

        private async Task LogoutAction()
        {
            await _authService.Logout();
            App.Current.Shutdown();
        }
    }
}