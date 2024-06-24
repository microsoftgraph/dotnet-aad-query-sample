using CommunityToolkit.Mvvm.DependencyInjection;
using CommunityToolkit.Mvvm.Messaging;
using CommunityToolkit.WinUI.UI.Controls;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Data;
using MsGraphSamples.WinUI.Converters;
using MsGraphSamples.WinUI.ViewModels;

// To learn more about WinUI, the WinUI project structure,
// and more about our project templates, see: http://aka.ms/winui-project-info.

namespace MsGraphSamples.WinUI.Views
{
    /// <summary>
    /// An empty page that can be used on its own or navigated to within a Frame.
    /// </summary>
    public sealed partial class MainPage : Page, IRecipient<string[]>
    {
        public MainViewModel? ViewModel { get; } = Ioc.Default.GetService<MainViewModel>();

        public MainPage()
        {
            this.InitializeComponent();
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

        private void TextBox_SelectAll(object sender, RoutedEventArgs _)
        {
            var textBox = (TextBox)sender;
            textBox.SelectAll();
        }
        private void Page_Loaded(object _, RoutedEventArgs __)
        {
            ViewModel?.Init().Await();
        }
    }
}
