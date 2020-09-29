// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using MsGraph_Samples.Helpers;
using MsGraph_Samples.Services;

namespace MsGraph_Samples.ViewModels
{
    public class ViewModelLocator
    {
        private readonly IAuthService _authService;
        private MainViewModel? _mainVM;
        public MainViewModel MainVM => _mainVM ??= new MainViewModel(_authService);

        public ViewModelLocator()
        {
            _authService = new AuthService(SecretConfig.ClientId);
        }
    }
}