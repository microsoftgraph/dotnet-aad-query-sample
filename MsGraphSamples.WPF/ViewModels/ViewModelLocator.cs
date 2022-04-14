// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using System.Diagnostics.CodeAnalysis;
using System.Windows;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Toolkit.Mvvm.DependencyInjection;
using MsGraph_Samples.Services;

namespace MsGraph_Samples.ViewModels
{
    [SuppressMessage("Performance", "CA1822:Mark members as static", Justification = "Binding Parameters")]
    public class ViewModelLocator
    {
        public static bool IsInDesignMode => Application.Current.MainWindow == null;

        public MainViewModel? MainVM => Ioc.Default.GetService<MainViewModel>();

        public ViewModelLocator()
        {
            Ioc.Default.ConfigureServices(GetServices());
        }

        private static IServiceProvider GetServices()
        {
            IServiceCollection serviceCollection = new ServiceCollection();

            if (IsInDesignMode)
            {
                serviceCollection.AddSingleton<IAuthService, FakeAuthService>();
                serviceCollection.AddSingleton<IGraphDataService, FakeGraphDataService>();
            }
            else
            {
                var authService = new AuthService();
                serviceCollection.AddSingleton<IAuthService>(authService);
                
                var graphDataService = new GraphDataService(authService.GraphClient);
                serviceCollection.AddSingleton<IGraphDataService>(graphDataService);
            }

            serviceCollection.AddSingleton<MainViewModel>();

            return serviceCollection.BuildServiceProvider();
        }
    }
}