// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using System.Diagnostics;
using System.Net;
using System.Text;
using System.Windows.Controls;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using CommunityToolkit.Mvvm.Messaging;
using Microsoft.Graph.Models;
using Microsoft.Graph.Models.ODataErrors;
using Microsoft.Kiota.Abstractions;
using MsGraphSamples.Services;
using MsGraphSamples.WPF.Helpers;


namespace MsGraphSamples.WPF.ViewModels;

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

    public static IReadOnlyList<string> Entities => new[] { "Users", "Groups", "Applications", "ServicePrincipals", "Devices" };

    [ObservableProperty]
    private string _selectedEntity = "Users";

    [ObservableProperty]
    [NotifyCanExecuteChangedFor(nameof(DrillDownCommand))]
    private DirectoryObject? _selectedObject;

    [ObservableProperty]
    [NotifyCanExecuteChangedFor(nameof(LaunchGraphExplorerCommand))]
    [NotifyCanExecuteChangedFor(nameof(LoadNextPageCommand))]
    private BaseCollectionPaginationCountResponse? _directoryObjects;

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

    public MainViewModel(IAuthService authService, IGraphDataService graphDataService)
    {
        _authService = authService;
        _graphDataService = graphDataService;

        Init().Await();
    }

    public async Task Init()
    {
        //var user = await _graphDataService.GetUserAsync(new[] { "displayName" });
        //UserName = user?.DisplayName;

        await Load();
    }


    [RelayCommand]
    private Task Load()
    {
        return IsBusyWrapper(async () => DirectoryObjects = SelectedEntity switch
        {
            "Users" => await _graphDataService.GetUserCollectionAsync(SplittedSelect, Filter, SplittedOrderBy, Search),
            "Groups" => await _graphDataService.GetGroupCollectionAsync(SplittedSelect, Filter, SplittedOrderBy, Search),
            "Applications" => await _graphDataService.GetApplicationCollectionAsync(SplittedSelect, Filter, SplittedOrderBy, Search),
            "ServicePrincipals" => await _graphDataService.GetServicePrincipalsCollectionAsync(SplittedSelect, Filter, SplittedOrderBy, Search),
            "Devices" => await _graphDataService.GetDeviceCollectionAsync(SplittedSelect, Filter, SplittedOrderBy, Search),
            _ => throw new NotImplementedException("Can't find selected entity")
        });
    }

    private bool CanDrillDown() => SelectedObject is not null;
    [RelayCommand(CanExecute = nameof(CanDrillDown))]
    private Task DrillDown()
    {
        ArgumentNullException.ThrowIfNull(SelectedObject);

        return IsBusyWrapper(async () =>
        {
            OrderBy = string.Empty;
            Filter = string.Empty;
            Search = string.Empty;

            DirectoryObjects = SelectedEntity switch
            {
                "Users" => await _graphDataService.GetTransitiveMemberOfAsGroupCollectionAsync(SelectedObject.Id!, SplittedSelect),
                "Groups" => await _graphDataService.GetTransitiveMembersAsUserCollectionAsync(SelectedObject.Id!, SplittedSelect),
                "Applications" => await _graphDataService.GetAppOwnersAsUserCollectionAsync(SelectedObject.Id!, SplittedSelect),
                "ServicePrincipals" => await _graphDataService.GetSPOwnersAsUserCollectionAsync(SelectedObject.Id!, SplittedSelect),
                "Devices" => await _graphDataService.GetRegisteredOwnersAsUserCollectionAsync(SelectedObject.Id!, SplittedSelect),
                _ => throw new NotImplementedException("Can't find selected entity")
            };
        });
    }

    private bool CanGoNextPage => DirectoryObjects?.OdataNextLink is not null;
    [RelayCommand(CanExecute = nameof(CanGoNextPage))]
    private async Task LoadNextPage()
    {
        DirectoryObjects = DirectoryObjects switch
        {
            UserCollectionResponse userCollection => await _graphDataService.GetNextPageAsync(userCollection),
            GroupCollectionResponse groupCollection => await _graphDataService.GetNextPageAsync(groupCollection),
            ApplicationCollectionResponse applicationCollection => await _graphDataService.GetNextPageAsync(applicationCollection),
            ServicePrincipalCollectionResponse servicePrincipalCollection => await _graphDataService.GetNextPageAsync(servicePrincipalCollection),
            DeviceCollectionResponse deviceCollection => await _graphDataService.GetNextPageAsync(deviceCollection),
            _ => throw new NotImplementedException("Can't find selected entity")
        };
    }

    [RelayCommand]
    private Task Sort(DataGridSortingEventArgs? e)
    {
        ArgumentNullException.ThrowIfNull(e);

        OrderBy = (string)e.Column.Header;
        e.Handled = true;
        return Load();
    }

    private bool CanLaunchGraphExplorer => LastUrl is not null;
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

        // Sending message to generate DataGridColumns according to the selected properties
        WeakReferenceMessenger.Default.Send(SplittedSelect);

        try
        {
            await task();

            SelectedEntity = DirectoryObjects switch
            {
                UserCollectionResponse => "Users",
                GroupCollectionResponse => "Groups",
                ApplicationCollectionResponse => "Applications",
                ServicePrincipalCollectionResponse => "ServicePrincipals",
                DeviceCollectionResponse => "Devices",
                _ => SelectedEntity,
            };
        }
        catch (ODataError ex)
        {
            Task.Run(() => System.Windows.MessageBox.Show(ex.Message, ex.Error?.Message)).Await();
        }
        catch (ApiException ex)
        {
            Task.Run(() => System.Windows.MessageBox.Show(ex.Message, ex.Source)).Await();
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