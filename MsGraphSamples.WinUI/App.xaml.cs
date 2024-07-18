using CommunityToolkit.Mvvm.DependencyInjection;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Navigation;
using MsGraphSamples.Services;
using MsGraphSamples.WinUI.Helpers;
using MsGraphSamples.WinUI.ViewModels;
using MsGraphSamples.WinUI.Views;
using Windows.ApplicationModel;

// To learn more about WinUI, the WinUI project structure,
// and more about our project templates, see: http://aka.ms/winui-project-info.

namespace MsGraphSamples.WinUI;


/// <summary>
/// Provides application-specific behavior to supplement the default Application class.
/// </summary>
public partial class App : Application
{
    /// <summary>
    /// Initializes the singleton application object.  This is the first line of authored code
    /// executed, and as such is the logical equivalent of main() or WinMain().
    /// </summary>
    public App()
    {
        this.InitializeComponent();
        Ioc.Default.ConfigureServices(GetServices());
    }

    private static ServiceProvider GetServices()
    {
        var serviceCollection = new ServiceCollection();

        if (!DesignMode.DesignModeEnabled)
        {
            var authService = new AuthService();
            serviceCollection.AddSingleton<IAuthService>(authService);

            var asyncEnumerableGraphDataService = new AsyncEnumerableGraphDataService(authService.GraphClient);
            serviceCollection.AddSingleton<IAsyncEnumerableGraphDataService>(asyncEnumerableGraphDataService);

            serviceCollection.AddSingleton<IDialogService, DialogService>();
        }

        serviceCollection.AddTransient<MainViewModel>();

        return serviceCollection.BuildServiceProvider();
    }


    /// <summary>
    /// Invoked when the application is launched.
    /// </summary>
    /// <param name="args">Details about the launch request and process.</param>
    protected override void OnLaunched(LaunchActivatedEventArgs args)
    {
        var mainWindow = new MainWindow { ExtendsContentIntoTitleBar = true };

        // Create a Frame to act as the navigation context and navigate to the first page
        var rootFrame = new Frame();
        rootFrame.Loaded += Root_Loaded;
        rootFrame.NavigationFailed += OnNavigationFailed;

        // Navigate to the first page, configuring the new page
        // by passing required information as a navigation parameter
        rootFrame.Navigate(typeof(MainPage), args.Arguments);

        // Place the frame in the current Window
        mainWindow.Content = rootFrame;

        // Ensure the MainWindow is active
        mainWindow.Activate();
    }

    private void Root_Loaded(object sender, RoutedEventArgs e)
    {
        // ContentDialog requires a reference to XamlRoot that is only present after Load.
        var dialogService = Ioc.Default.GetRequiredService<IDialogService>();
        dialogService.Root = ((Frame)sender).XamlRoot;
    }

    void OnNavigationFailed(object sender, NavigationFailedEventArgs e)
    {
        throw new Exception("Failed to load Page " + e.SourcePageType.FullName);
    }
}
