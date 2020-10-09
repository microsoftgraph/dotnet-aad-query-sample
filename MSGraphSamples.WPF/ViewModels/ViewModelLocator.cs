// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using System;
using System.Windows;
using Microsoft.Extensions.DependencyInjection;
using MsGraph_Samples.Helpers;
using MsGraph_Samples.Services;

namespace MsGraph_Samples.ViewModels
{
    public class ViewModelLocator
    {
        public static bool IsInDesignMode => Application.Current.MainWindow == null;

        public IServiceProvider Services { get; }

        public MainViewModel MainVM => Services.GetService<MainViewModel>();

        public ViewModelLocator()
        {
            IServiceCollection serviceCollection = new ServiceCollection();
            ConfigureServices(serviceCollection);
            Services = serviceCollection.BuildServiceProvider();
        }

        private void ConfigureServices(IServiceCollection serviceCollection)
        {
            if(IsInDesignMode)
            {
                serviceCollection.AddSingleton<IAuthService, FakeAuthService>();
                serviceCollection.AddSingleton<IGraphDataService, FakeGraphDataService>();
            }
            else
            {
                var authService = new AuthService(SecretConfig.ClientId);
                serviceCollection.AddSingleton<IAuthService>(authService);

                var graphDataService = new GraphDataService(authService.GetServiceClient());
                serviceCollection.AddSingleton<IGraphDataService>(graphDataService);
            }

            serviceCollection.AddSingleton<MainViewModel>();
        }
    }
}