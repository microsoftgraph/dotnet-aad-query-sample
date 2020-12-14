// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using System.Linq;
using System.Threading.Tasks;
using Microsoft.Graph;
using Microsoft.Identity.Client;
using MsGraph_Samples.Helpers;

namespace MsGraph_Samples.Services
{
    public interface IAuthService 
    {
        GraphServiceClient GetServiceClient();
        Task<IAccount?> GetAccount();
        Task Logout();
    }

    public class AuthService : IAuthService
    {
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
        private GraphServiceClient? _graphClient;

        public AuthService(string clientId)
        {
            _publicClientApp = PublicClientApplicationBuilder.Create(clientId)
                .WithAuthority(CloudInstance, Tenant)
                .WithDefaultRedirectUri()
                .Build();

            TokenCacheHelper.EnableSerialization(_publicClientApp.UserTokenCache);
        }

        public GraphServiceClient GetServiceClient()
        {
            var authenticationProvider = new InteractiveAuthenticationProvider(_publicClientApp, _scopes);
            return _graphClient ??= new GraphServiceClient(authenticationProvider);
        }

        public async Task<IAccount?> GetAccount()
        {
            var accounts = await _publicClientApp.GetAccountsAsync();
            return accounts.FirstOrDefault();
        }

        public async Task Logout()
        {
            _graphClient = null;
            TokenCacheHelper.Clear();

            var account = await GetAccount();
            await _publicClientApp.RemoveAsync(account);
        }
    }
}