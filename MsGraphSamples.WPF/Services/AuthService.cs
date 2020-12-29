// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using Microsoft.Graph;
using Microsoft.Graph.Auth;
using Microsoft.Identity.Client;
using Microsoft.Identity.Client.Extensions.Msal;

namespace MsGraph_Samples.Services
{
    public interface IAuthService
    {
        IGraphServiceClient GetServiceClient();
        Task Logout();
    }

    public class AuthService : IAuthService
    {
        private static readonly string LocalAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
        private static readonly string ProjectName = Assembly.GetCallingAssembly().GetName().Name ?? "tokencache";
        private static readonly string CacheDirectoryPath = Path.Combine(LocalAppData, ProjectName);
        private const string CacheFileName = "msalcache.bin";

        /// <summary>
        /// The content of Tenant by the information about the accounts allowed to sign-in in your application:
        /// - for Work or School account in your org, use your tenant ID, or domain
        /// - for any Work or School accounts, use organizations
        /// - for any Work or School accounts, or Microsoft personal account, use common
        /// - for Microsoft Personal account, use consumers
        /// </summary>
        private const string Tenant = "organizations";

        // To change from Microsoft public cloud to a national cloud, use another value of AzureCloudInstance
        private const AzureCloudInstance CloudInstance = AzureCloudInstance.AzurePublic;

        // Make sure the user you login with has "Directory.Read.All" permissions
        private readonly string[] _scopes = { "Directory.Read.All" };

        private readonly IPublicClientApplication _publicClientApp;

        private InteractiveAuthenticationProvider AuthProvider => new(_publicClientApp, _scopes);
        
        private IGraphServiceClient? _graphClient;
        public IGraphServiceClient GetServiceClient() => _graphClient ??= new GraphServiceClient(AuthProvider);

        public AuthService(string clientId)
        {
            _publicClientApp = PublicClientApplicationBuilder.Create(clientId)
                .WithAuthority(CloudInstance, Tenant)
                .WithDefaultRedirectUri()
                .Build();

            var storageCreationProperties = new StorageCreationPropertiesBuilder(CacheFileName, CacheDirectoryPath, clientId).Build();

            // Workaround for Async creation, waiting for
            // https://github.com/AzureAD/microsoft-authentication-extensions-for-dotnet/issues/102
            MsalCacheHelper
                .CreateAsync(storageCreationProperties)
                .Await(ch => ch.RegisterCache(_publicClientApp.UserTokenCache));
        }

        public async Task Logout()
        {
            _graphClient = null;

            var account = await GetAccount();
            await _publicClientApp.RemoveAsync(account);
        }

        private async Task<IAccount?> GetAccount()
        {
            var accounts = await _publicClientApp.GetAccountsAsync();
            return accounts.FirstOrDefault();
        }
    }
}