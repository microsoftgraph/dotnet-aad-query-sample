// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using System;
using System.Diagnostics;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using Microsoft.Graph;
using Microsoft.Identity.Client;
using Microsoft.Identity.Client.Extensions.Msal;
using System.Reflection;

namespace MsGraph_Samples.Services
{
    public interface IAuthService : IAuthenticationProvider
    {
        IAccount? Account { get; }
        event Action? AuthenticationSuccessful;
        GraphServiceClient GetServiceClient();
        Task Logout();
    }

    public class AuthService : IAuthService
    {
        public IAccount? Account { get; private set; }
        public event Action? AuthenticationSuccessful;

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
        private readonly MsalCacheHelper _cacheHelper;
        private GraphServiceClient? _graphClient;

        public AuthService(string clientId)
        {
            _publicClientApp = PublicClientApplicationBuilder.Create(clientId)
                .WithAuthority(CloudInstance, Tenant)
                .WithDefaultRedirectUri()
                .Build();

            var LocalAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            var ProjectName = Assembly.GetCallingAssembly().GetName().Name ?? "tokencache";
            var CacheFilePath = $"{LocalAppData}\\{ProjectName}\\";

            var storageCreationProperties = new StorageCreationPropertiesBuilder("msalcache.bin", CacheFilePath, clientId).Build();

            // .Result is not a best practice, waiting for https://github.com/AzureAD/microsoft-authentication-extensions-for-dotnet/issues/102
            _cacheHelper = MsalCacheHelper.CreateAsync(storageCreationProperties).Result;
            _cacheHelper.RegisterCache(_publicClientApp.UserTokenCache);
        }

        public GraphServiceClient GetServiceClient()
        {
            var authProvider = new DelegateAuthenticationProvider(AuthenticateRequestAsync);
            return _graphClient ??= new GraphServiceClient(authProvider);
        }

        public async Task AuthenticateRequestAsync(HttpRequestMessage requestMessage)
        {
            var accessToken = await AcquireTokenAsync();
            requestMessage.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
            if (accessToken != null)
                AuthenticationSuccessful?.Invoke();
        }

        private async Task<string?> AcquireTokenAsync()
        {
            var accounts = await _publicClientApp
                .GetAccountsAsync().ConfigureAwait(false);
            Account = accounts.FirstOrDefault();

            AuthenticationResult authResult;

            try //Trying to acquire token silently
            {
                authResult = await _publicClientApp
                    .AcquireTokenSilent(_scopes, Account)
                    .ExecuteAsync().ConfigureAwait(false);
            }
            catch (MsalUiRequiredException ex1)
            {
                Debug.WriteLine($"MsalUiRequiredException: {ex1.Message}");
                try //Trying to acquire token using Integrated Windows Auth
                {
                    authResult = await _publicClientApp
                        .AcquireTokenByIntegratedWindowsAuth(_scopes)
                        .ExecuteAsync().ConfigureAwait(false);
                }
                catch (MsalException ex2)
                {
                    Debug.WriteLine($"MsalClientException: {ex2.Message}");
                    try //Trying to acquire token via Web page
                    {
                        authResult = await _publicClientApp
                            .AcquireTokenInteractive(_scopes)
                            .WithClaims(ex1.Claims)
                            .ExecuteAsync().ConfigureAwait(false);

                        Account = authResult.Account;
                    }
                    catch (MsalException ex3)
                    {
                        Debug.WriteLine($"Error Acquiring Token:{Environment.NewLine}{ex3}");
                        return null;
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error Acquiring Token Silently:{Environment.NewLine}{ex}");
                return null;
            }

            return authResult?.AccessToken;
        }

        public async Task Logout()
        {
            _cacheHelper.Clear();
            _graphClient = null;
            await _publicClientApp.RemoveAsync(Account).ConfigureAwait(false);
        }
    }
}