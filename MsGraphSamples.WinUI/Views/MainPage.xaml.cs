using CommunityToolkit.Mvvm.DependencyInjection;
using CommunityToolkit.Mvvm.Messaging;
using CommunityToolkit.WinUI.UI.Controls;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Data;
using MsGraphSamples.WinUI.Converters;
using MsGraphSamples.WinUI.Helpers;
using MsGraphSamples.WinUI.ViewModels;
using System.Collections.Immutable;
using System.Reflection;
using Windows.Foundation;

namespace MsGraphSamples.WinUI.Views;
public sealed partial class MainPage : Page, IRecipient<ImmutableSortedDictionary<string, DataGridSortDirection?>>
{
    public MainViewModel ViewModel { get; } = Ioc.Default.GetRequiredService<MainViewModel>();
    private readonly Debouncer debouncer = new(TimeSpan.FromMilliseconds(600));

    public MainPage()
    {
        this.InitializeComponent();
        WeakReferenceMessenger.Default.Register(this);
    }

    public void Receive(ImmutableSortedDictionary<string, DataGridSortDirection?> properties)
    {
        DirectoryObjectsGrid.Columns.Clear();

        foreach (var property in properties)
        {
            // handle extension properties
            if (property.Key.StartsWith("extension_"))
            {
                DirectoryObjectsGrid.Columns.Add(new DataGridTextColumn
                {
                    Header = property.Key.Split('_')[2],
                    Binding = new Binding() { Path = new PropertyPath("AdditionalData"), Converter = new AdditionalDataConverter(), ConverterParameter = property.Key },
                    SortDirection = property.Value,
                    Width = new DataGridLength(1, DataGridLengthUnitType.Star)
                });
            }
            else
            {
                DirectoryObjectsGrid.Columns.Add(new DataGridTextColumn
                {
                    Header = property.Key,
                    Binding = new Binding() { Path = new PropertyPath(property.Key), Converter = new CollectionToCommaSeparatedStringConverter() },
                    SortDirection = property.Value,
                    Width = new DataGridLength(1, DataGridLengthUnitType.Star)
                });
            }
        }
    }

    private void TextBox_SelectAll(object sender, RoutedEventArgs _)
    {
        var textBox = (TextBox)sender;
        textBox.SelectAll();
    }

    private void DirectoryObjectsGrid_SizeChanged(object sender, SizeChangedEventArgs e)
    {
        debouncer.Debounce(SetPageSize);
    }

    private void SetPageSize()
    {
        var rowsPresenterAvailableSizeObj = typeof(DataGrid)
            .GetField("_rowsPresenterAvailableSize", BindingFlags.NonPublic | BindingFlags.Instance)
            ?.GetValue(DirectoryObjectsGrid);

        var rowHeightEstimateObj = typeof(DataGrid)
            .GetProperty("RowHeightEstimate", BindingFlags.NonPublic | BindingFlags.Instance)
            ?.GetValue(DirectoryObjectsGrid);

        if (rowsPresenterAvailableSizeObj is Size rowPresenterSize &&
            rowHeightEstimateObj is double rowHeight)
        {
            ViewModel.PageSize = (ushort)Math.Min(Math.Round(DirectoryObjectsGrid.DataFetchSize * rowPresenterSize.Height / rowHeight), 999);
        }
    }
}