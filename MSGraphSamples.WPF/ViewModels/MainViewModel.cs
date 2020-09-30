// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Net;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Graph;
using MsGraph_Samples.MVVM;
using MsGraph_Samples.Services;

namespace MsGraph_Samples.ViewModels
{
    public class MainViewModel : ViewModelBase
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

        public string? UserName => _authService.Account?.Username;
        public string? LastUrl => _graphDataService.LastUrl;

        public IReadOnlyList<string> Entities => new[] { "Users", "Groups", "Applications", "Devices" };
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
            set => Set(ref _filter,value);
        }

        private string _search = string.Empty;
        public string Search
        {
            get => _search;
            set => Set(ref _search, value);
        }

        private string _orderBy = "displayName";
        public string OrderBy
        {
            get => _orderBy;
            set => Set(ref _orderBy, value);
        }

        public MainViewModel(IAuthService authService)
        {
            _authService = authService;
            _authService.AuthenticationSuccessful += () =>
            {
                RaisePropertyChanged(nameof(UserName));
                LogoutCommand.RaiseCanExecuteChanged();
            };

            if (IsInDesignMode)
            {
                _graphDataService = new FakeGraphDataService();
                return;
            }

            _graphDataService = new GraphDataService(_authService.GetServiceClient());
            LoadAction();
        }

        public RelayCommand LoadCommand => new RelayCommand(LoadAction);
        private async void LoadAction()
        {
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
                    _ => throw new NotImplementedException("Can't find selected entity"),
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
                GraphExplorerCommand.RaiseCanExecuteChanged();
                IsBusy = false;
            }
        }

        private RelayCommand? _drillDownCommand;
        public RelayCommand DrillDownCommand => _drillDownCommand ??= new RelayCommand(DrillDownAction);
        private async void DrillDownAction()
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
                GraphExplorerCommand.RaiseCanExecuteChanged();
                IsBusy = false;
            }
        }

        public RelayCommand<DataGridAutoGeneratingColumnEventArgs> AutoGeneratingColumn =>
            new RelayCommand<DataGridAutoGeneratingColumnEventArgs>(AutoGeneratingColumnAction);
        private void AutoGeneratingColumnAction(DataGridAutoGeneratingColumnEventArgs e)
        {
            if (!string.IsNullOrEmpty(Select))
            {
                e.Cancel = !e.PropertyName.In(Select.Split(','));
                e.Column.Width = new DataGridLength(1, DataGridLengthUnitType.Star);
            }
        }

        private RelayCommand<DataGridSortingEventArgs>? _sortCommand;
        public RelayCommand<DataGridSortingEventArgs> SortCommand => _sortCommand ??= new RelayCommand<DataGridSortingEventArgs>(SortAction);
        private void SortAction(DataGridSortingEventArgs e)
        {
            OrderBy = $"{e.Column.Header}";
            e.Handled = true;
            LoadAction();
        }

        private RelayCommand? _graphExplorerCommand;
        public RelayCommand GraphExplorerCommand => _graphExplorerCommand ??= new RelayCommand(GraphExplorerAction, () => LastUrl != null);
        private void GraphExplorerAction()
        {
            if (LastUrl == null) return;

            var geBaseUrl = "https://developer.microsoft.com/en-us/graph/graph-explorer";
            var version = "v1.0";
            var graphUrl = "https://graph.microsoft.com";
            var encodedUrl = WebUtility.UrlEncode(LastUrl.Substring(LastUrl.NthIndexOf('/', 4) + 1));
            var encodedHeaders = "W3sibmFtZSI6IkNvbnNpc3RlbmN5TGV2ZWwiLCJ2YWx1ZSI6ImV2ZW50dWFsIn1d";
            var url = $"{geBaseUrl}?request={encodedUrl}&method=GET&version={version}&GraphUrl={graphUrl}&headers={encodedHeaders}";

            var psi = new ProcessStartInfo { FileName = url, UseShellExecute = true };
            System.Diagnostics.Process.Start(psi);
        }

        private RelayCommand? _logoutCommand;
        public RelayCommand LogoutCommand => _logoutCommand ??= new RelayCommand(LogoutAction, () => UserName != null);
        private void LogoutAction()
        {
            _authService.Logout();
            App.Current.Shutdown();
        }
    }
}