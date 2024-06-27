// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using System.Diagnostics.CodeAnalysis;
using System.Windows;
using CommunityToolkit.Mvvm.DependencyInjection;
using Microsoft.Extensions.DependencyInjection;
using MsGraphSamples.Services;

namespace MsGraphSamples.WPF.ViewModels;

[SuppressMessage("Performance", "CA1822:Mark members as static", Justification = "Binding Parameters")]
public class ViewModelLocator
{
    public static bool IsInDesignMode => Application.Current.MainWindow == null;

    public MainViewModel? MainVM => Ioc.Default.GetService<MainViewModel>();

    public ViewModelLocator()
    {
        Ioc.Default.ConfigureServices(GetServices());
    }

    private static ServiceProvider GetServices()
    {
        var serviceCollection = new ServiceCollection();

        if (!IsInDesignMode)
        {
            var authService = new AuthService();
            serviceCollection.AddSingleton<IAuthService>(authService);

            var graphDataService = new GraphDataService(authService.GraphClient);
            serviceCollection.AddSingleton<IGraphDataService>(graphDataService);
        }

        serviceCollection.AddTransient<MainViewModel>();

        return serviceCollection.BuildServiceProvider();
    }
}