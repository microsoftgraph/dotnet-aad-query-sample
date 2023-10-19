﻿using System.Diagnostics;
using System.Net;
using System.Text;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using CommunityToolkit.Mvvm.Messaging;
using CommunityToolkit.WinUI.UI.Controls;
using Microsoft.Graph.Models;
using Microsoft.Graph.Models.ODataErrors;
using Microsoft.Kiota.Abstractions;
using MsGraphSamples.Services;
using MsGraphSamples.WinUI.Contracts.ViewModels;
using MsGraphSamples.WinUI.Helpers;

namespace MsGraphSamples.WinUI.ViewModels;

public partial class MainViewModel : ObservableRecipient, INavigationAware
{
    private readonly IAsyncEnumerableGraphDataService _graphDataService;

    private readonly ushort pageSize = 25;

    private readonly Stopwatch _stopWatch = new();
    public long ElapsedMs => _stopWatch.ElapsedMilliseconds;

    [ObservableProperty]
    private bool _isBusy = false;

    [ObservableProperty]
    private bool _isError = false;

    [ObservableProperty]
    private string? _userName;

    [ObservableProperty]
    [NotifyCanExecuteChangedFor(nameof(LaunchGraphExplorerCommand))]
    private AsyncLoadingCollection<DirectoryObject>? _directoryObjects;

    [ObservableProperty]
    [NotifyCanExecuteChangedFor(nameof(DrillDownCommand))]
    private DirectoryObject? _selectedObject;

    public IReadOnlyList<string> Entities => new[] { "Users", "Groups", "Applications", "ServicePrincipals", "Devices" };

    [ObservableProperty]
    private string _selectedEntity = "Users";
    public string? LastUrl => _graphDataService.LastUrl;
    public long? LastCount => _graphDataService.LastCount;

    #region OData Operators

    public string[] SplittedSelect => Select.Split(',', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries);

    [ObservableProperty]
    public string _select = "id,displayName,mail,userPrincipalName";

    [ObservableProperty]
    public string? _filter = string.Empty;

    public string[]? SplittedOrderBy => OrderBy?.Split(',', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries);

    [ObservableProperty]
    public string? _orderBy;

    private string? _search;
    public string? Search
    {
        get => _search;
        set
        {
            if (_search != value)
            {
                _search = FixSearchSyntax(value);
                OnPropertyChanged();
            }
        }
    }

    private static string? FixSearchSyntax(string? searchValue)
    {
        if (searchValue == null)
            return null;

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

    public MainViewModel(IAsyncEnumerableGraphDataService graphDataService)
    {
        _graphDataService = graphDataService;
    }

    public async void OnNavigatedTo(object parameter)
    {
        //var user = await _graphDataService.GetUserAsync(new[] { "displayName" });
        //UserName = user?.DisplayName;

        await Load();
    }

    public void OnNavigatedFrom()
    {
    }

    [RelayCommand]
    private Task Load()
    {
        return IsBusyWrapper(() => SelectedEntity switch
        {
            //"Users" =>  _graphDataService.GetUsersInBatch(SplittedSelect, pageSize),
            "Users" => _graphDataService.GetUsers(SplittedSelect, Filter, SplittedOrderBy, Search, pageSize),
            "Groups" => _graphDataService.GetGroups(SplittedSelect, Filter, SplittedOrderBy, Search, pageSize),
            "Applications" => _graphDataService.GetApplications(SplittedSelect, Filter, SplittedOrderBy, Search, pageSize),
            "ServicePrincipals" => _graphDataService.GetServicePrincipals(SplittedSelect, Filter, SplittedOrderBy, Search, pageSize),
            "Devices" => _graphDataService.GetDevices(SplittedSelect, Filter, SplittedOrderBy, Search, pageSize),
            _ => throw new NotImplementedException("Can't find selected entity")
        });
    }

    private bool CanDrillDown() => SelectedObject is not null;
    [RelayCommand(CanExecute = nameof(CanDrillDown))]
    private Task DrillDown()
    {
        ArgumentNullException.ThrowIfNull(SelectedObject);

        // Need to save the Id because the SelectedObject will be cleared by DirectoryObjects.Clear() inside IsBusyWrapper()
        var id = SelectedObject.Id!;

        return IsBusyWrapper(() =>
        {
            OrderBy = string.Empty;
            Filter = string.Empty;
            Search = string.Empty;

            return SelectedEntity switch
            {
                "Users" => _graphDataService.GetTransitiveMemberOfAsGroups(id, SplittedSelect, pageSize),
                "Groups" => _graphDataService.GetTransitiveMembersAsUsers(id, SplittedSelect, pageSize),
                "Applications" => _graphDataService.GetAppOwnersAsUsers(id, SplittedSelect, pageSize),
                "ServicePrincipals" => _graphDataService.GetSPOwnersAsUsers(id, SplittedSelect, pageSize),
                "Devices" => _graphDataService.GetRegisteredOwnersAsUsers(id, SplittedSelect, pageSize),
                _ => throw new NotImplementedException("Can't find selected entity")
            };
        });
    }


    [RelayCommand]
    private Task Sort(DataGridColumnEventArgs e)
    {
        ArgumentNullException.ThrowIfNull(e);

        OrderBy = (string)e.Column.Header;
        //e.handled  = true;
        return Load();
    }

    private bool CanLaunchGraphExplorer() => LastUrl is not null;
    [RelayCommand(CanExecute = nameof(CanLaunchGraphExplorer))]
    private void LaunchGraphExplorer()
    {
        ArgumentNullException.ThrowIfNull(LastUrl);

        var geBaseUrl = "https://developer.microsoft.com/en-us/graph/graph-explorer";
        var graphUrl = "https://graph.microsoft.com";
        var version = "v1.0";
        var startOfQuery = LastUrl.NthIndexOf('/', 4) + 1;
        var encodedUrl = WebUtility.UrlEncode(LastUrl[startOfQuery..]);
        var encodedHeaders = "W3sibmFtZSI6IkNvbnNpc3RlbmN5TGV2ZWwiLCJ2YWx1ZSI6ImV2ZW50dWFsIn1d"; // ConsistencyLevel = eventual

        var url = $"{geBaseUrl}?request={encodedUrl}&method=GET&version={version}&GraphUrl={graphUrl}&headers={encodedHeaders}";

        var psi = new ProcessStartInfo { FileName = url, UseShellExecute = true };
        System.Diagnostics.Process.Start(psi);
    }

    private async Task IsBusyWrapper(Func<IAsyncEnumerable<DirectoryObject>> getDirectoryObjects)
    {
        IsError = false;
        IsBusy = true;
        _stopWatch.Restart();

        // Sending message to generate DataGridColumns according to the selected properties
        WeakReferenceMessenger.Default.Send(SplittedSelect);

        try
        {
            DirectoryObjects = new(getDirectoryObjects(), pageSize);
            await DirectoryObjects.LoadMoreItemsAsync();
            
            SelectedEntity = DirectoryObjects.FirstOrDefault() switch
            {
                User => "Users",
                Group => "Groups",
                Application => "Applications",
                ServicePrincipal => "ServicePrincipals",
                Device => "Devices",
                _ => SelectedEntity,
            };
        }
        catch (ODataError ex)
        {
            IsError = true;
            await App.MainWindow.ShowMessageDialogAsync(ex.Message, ex.Error?.Message ?? string.Empty);
        }
        catch (ApiException ex)
        {
            IsError = true;
            await App.MainWindow.ShowMessageDialogAsync(ex.Message, ex.Source ?? string.Empty);
        }
        finally
        {
            _stopWatch.Stop();
            OnPropertyChanged(nameof(ElapsedMs));
            OnPropertyChanged(nameof(LastUrl));
            OnPropertyChanged(nameof(LastCount));
            IsBusy = false;
        }
    }
}