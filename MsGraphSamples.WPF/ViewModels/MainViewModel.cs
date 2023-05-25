// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using System.Collections.ObjectModel;
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
    private readonly IAsyncEnumerableGraphDataService _graphDataService;

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
    private ObservableCollection<DirectoryObject> _directoryObjects = new();

    #region OData Operators

    public string[] SplittedSelect => Select.Split(',', StringSplitOptions.TrimEntries);

    [ObservableProperty]
    public string _select = "id,displayName,mail,userPrincipalName";

    [ObservableProperty]
    public string? _filter = string.Empty;

    public string[]? SplittedOrderBy => OrderBy?.Split(',', StringSplitOptions.TrimEntries);

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

    public MainViewModel(IAuthService authService, IAsyncEnumerableGraphDataService graphDataService)
    {
        _authService = authService;
        _graphDataService = graphDataService;

        Init().Await();
    }

    public async Task Init()
    {
        var user = await _graphDataService.GetUserAsync(new[] { "displayName" });
        UserName = user?.DisplayName;

        await Load();
    }


    [RelayCommand]
    private Task Load()
    {
        return IsBusyWrapper(() => SelectedEntity switch
        {
            "Users" => _graphDataService.GetUsers(SplittedSelect, Filter, SplittedOrderBy, Search),
            "Groups" => _graphDataService.GetGroups(SplittedSelect, Filter, SplittedOrderBy, Search),
            "Applications" => _graphDataService.GetApplications(SplittedSelect, Filter, SplittedOrderBy, Search),
            "Devices" => _graphDataService.GetDevices(SplittedSelect, Filter, SplittedOrderBy, Search),
            _ => throw new NotImplementedException("Can't find selected entity")
        });
    }

    private bool CanDrillDown() => SelectedObject is not null;
    [RelayCommand(CanExecute = nameof(CanDrillDown))]
    private Task DrillDown()
    {
        ArgumentNullException.ThrowIfNull(SelectedObject);

        return IsBusyWrapper(() =>
        {
            OrderBy = string.Empty;
            Filter = string.Empty;
            Search = string.Empty;

            return SelectedEntity switch
            {
                "Users" => _graphDataService.GetTransitiveMemberOfAsGroups(SelectedObject.Id!),
                "Groups" => _graphDataService.GetTransitiveMembersAsUsers(SelectedObject.Id!),
                "Applications" => _graphDataService.GetAppOwnersAsUsers(SelectedObject.Id!),
                "Devices" => _graphDataService.GetTransitiveMemberOfAsGroups(SelectedObject.Id!),
                _ => throw new NotImplementedException("Can't find selected entity")
            };
        });
    }

    [RelayCommand]
    private Task Sort(DataGridSortingEventArgs? e)
    {
        ArgumentNullException.ThrowIfNull(e);

        OrderBy = (string)e.Column.Header;
        e.Handled = true;
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

    [RelayCommand]
    private void Logout()
    {
        _authService.Logout();
        App.Current.Shutdown();
    }

    private async Task IsBusyWrapper(Func<IAsyncEnumerable<DirectoryObject>> getDirectoryObjects)
    {
        IsBusy = true;
        _stopWatch.Restart();
        DirectoryObjects.Clear();

        // Sending message to generate DataGridColumns according to the selected properties
        WeakReferenceMessenger.Default.Send(SplittedSelect);

        try
        {
            await foreach (var item in getDirectoryObjects())
            {
                DirectoryObjects.Add(item);
            }

            SelectedEntity = DirectoryObjects.FirstOrDefault() switch
            {
                User => "Users",
                Group => "Groups",
                Application => "Applications",
                Device => "Devices",
                _ => SelectedEntity,
            };

            LaunchGraphExplorerCommand.NotifyCanExecuteChanged();
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