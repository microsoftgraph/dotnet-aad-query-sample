// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Graph;
using MsGraph_Samples.Helpers;
using MsGraph_Samples.Services;

namespace MsGraph_Samples.ViewModels
{
    public class MainViewModel : Observable
    {
        private readonly IGraphDataService _graphDataService;

        private bool _isBusy;
        public bool IsBusy
        {
            get => _isBusy;
            set => Set(ref _isBusy, value);
        }
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

        private IEnumerable<DirectoryObject>? _directoryObjects; // = Enumerable.Empty<DirectoryObject>();
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

        public string Select { get; set; } = "id, displayName, mail, userPrincipalName";

        public string Filter { get; set; } = string.Empty;

        public string Search { get; set; } = string.Empty;

        private string _orderBy = "displayName";
        public string OrderBy
        {
            get => _orderBy;
            set => Set(ref _orderBy, value);
        }

        public MainViewModel(IGraphDataService dataService)
        {
            _graphDataService = dataService;
            LoadAction();
        }

        public RelayCommand LoadCommand => new RelayCommand(LoadAction);

        private async void LoadAction()
        {
            IsBusy = true;

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
                var url = _graphDataService.LastCall;
                MessageBox.Show($"URL: {url}\n{ex.Message}", ex.Error.Message);
                Clipboard.SetText(url?.AbsoluteUri);
            }

            IsBusy = false;
        }

        public RelayCommand<DataGridAutoGeneratingColumnEventArgs> AutoGeneratingColumn =>
            new RelayCommand<DataGridAutoGeneratingColumnEventArgs>((e) => e.Cancel = !e.PropertyName.In(Select.Split(',')));

        public RelayCommand<DataGridSortingEventArgs> SortCommand => new RelayCommand<DataGridSortingEventArgs>(SortAction);
        private void SortAction(DataGridSortingEventArgs e)
        {
            OrderBy = $"{e.Column.Header}";
            e.Handled = true;
            LoadAction();
        }

        public RelayCommand DrillDownCommand => new RelayCommand(DrillDownCommandAction);
        private async void DrillDownCommandAction()
        {
            if (SelectedObject == null)
                return;

            Filter = string.Empty;

            try
            {
                DirectoryObjects = SelectedEntity switch
                {
                    "Users" => await _graphDataService.GetTransitiveMemberOfAsGroupsAsync(SelectedObject.Id),
                    "Groups" => await _graphDataService.GetTransitiveMembersAsUsersAsync(SelectedObject.Id),
                    "Applications" => await _graphDataService.GetOwnersAsUsersAsync(SelectedObject.Id),
                    "Devices" => await _graphDataService.GetTransitiveMemberOfAsGroupsAsync(SelectedObject.Id),
                    _ => null
                };
            }
            catch (ServiceException ex)
            {
                var url = _graphDataService.LastCall;
                MessageBox.Show($"URL: {url}\n{ex.Message}", ex.Error.Message);
                Clipboard.SetText(url?.AbsoluteUri);
            }
        }
    }
}