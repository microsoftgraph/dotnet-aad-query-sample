using CommunityToolkit.Mvvm.Messaging;
using CommunityToolkit.WinUI.UI.Controls;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Data;
using MsGraphSamples.WinUI.Helpers;
using MsGraphSamples.WinUI.ViewModels;

namespace MsGraphSamples.WinUI.Views;

public sealed partial class MainPage : Page, IRecipient<string[]>
{
    public MainViewModel ViewModel
    {
        get;
    }

    public MainPage()
    {
        ViewModel = App.GetService<MainViewModel>();
        InitializeComponent();
        WeakReferenceMessenger.Default.Register(this);
    }

    // Generate DataGridColumns
    public void Receive(string[] splittedSelect)
    {
        DirectoryObjectsGrid.Columns.Clear();

        foreach (var property in splittedSelect)
        {
            // handle extension properties
            if (property.StartsWith("extension_"))
            {
                DirectoryObjectsGrid.Columns.Add(new DataGridTextColumn
                {
                    Header = property.Split('_')[2],
                    Binding = new Binding() { Path = new PropertyPath("AdditionalData"), Converter = new AdditionalDataConverter(), ConverterParameter = property },
                    Width = new DataGridLength(1, DataGridLengthUnitType.Star)
                });
            }
            else
            {
                // TODO: find a more robust way to generate bindings with property names
                DirectoryObjectsGrid.Columns.Add(new DataGridTextColumn
                {
                    Header = property,
                    // binding needs exact casing of the property name (e.g. "UserPrincipalName" instead of "userPrincipalName")
                    Binding = new Binding() { Path = new PropertyPath(char.ToUpper(property[0]) + property[1..]) },
                    Width = new DataGridLength(1, DataGridLengthUnitType.Star)
                });
            }
        }
    }

    private void TextBox_SelectAll(object sender, RoutedEventArgs e)
    {
        var textBox = (TextBox)sender;
        textBox.SelectAll();
    }
}
