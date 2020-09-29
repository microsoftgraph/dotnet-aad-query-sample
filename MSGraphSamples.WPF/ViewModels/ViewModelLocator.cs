using System.IO;
using System.Windows;
using Microsoft.Extensions.Configuration;
using MsGraph_Samples.Services;

namespace MsGraph_Samples.ViewModels
{
    public class ViewModelLocator
    {
        private readonly IGraphDataService _graphDataService;

        private MainViewModel? _mainVM;
        public MainViewModel MainVM => _mainVM ??= new MainViewModel(_graphDataService);

        public ViewModelLocator()
        {
            if (IsInDesignMode)
            {
                _graphDataService = new FakeGraphDataService();
                return;
            }

            var appConfig = new ConfigurationBuilder()
                .AddUserSecrets<App>().Build();

            // This should contain your Client Id
            var clientId = appConfig["clientId"];
            if (clientId == null)
            {
                var helpUrl = "https://docs.microsoft.com/en-us/aspnet/core/security/app-secrets?view=aspnetcore-3.1&tabs=windows";
                throw new FileNotFoundException($"Missing or invalid secrets.json\nMake sure you created one: {helpUrl}");
            }

            var authService = new AuthService(clientId);
            var serviceClient = authService.GetServiceClient();
            _graphDataService = new GraphDataService(serviceClient);
        }

        private bool IsInDesignMode => Application.Current.MainWindow == null;
    }
}