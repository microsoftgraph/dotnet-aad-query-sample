// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using Azure.Core;
using Azure.Identity;
using Microsoft.Extensions.Configuration;
using Microsoft.Graph;

namespace MsGraphSamples.Services;

public interface IAuthService
{
    GraphServiceClient GraphClient { get; }
    void Logout();
}

public class AuthService : IAuthService
{
    private static readonly IConfiguration _configuration = new ConfigurationBuilder().AddUserSecrets<AuthService>().Build();

    private const string _tokenPath = "authToken.bin";
    private static readonly string[] _scopes = ["Directory.Read.All"];

    private GraphServiceClient? _graphClient;
    
    //public GraphServiceClient GraphClient => _graphClient ??= new GraphServiceClient(GetAppCredential());
    public GraphServiceClient GraphClient => _graphClient ??= new GraphServiceClient(GetBrowserCredential());
    
    private static ClientSecretCredential GetAppCredential() => new(
        _configuration["tenantId"],
        _configuration["clientId"],
        _configuration["clientSecret"]);

    private static InteractiveBrowserCredential GetBrowserCredential()
    {
        var credentialOptions = new InteractiveBrowserCredentialOptions
        {
            ClientId = _configuration["clientId"],
            TokenCachePersistenceOptions = new TokenCachePersistenceOptions() { UnsafeAllowUnencryptedStorage = true }
        };

        if (File.Exists(_tokenPath))
        {
            // use the cached token
            using var authRecordStream = File.OpenRead(_tokenPath);
            var authRecord = AuthenticationRecord.Deserialize(authRecordStream);
            credentialOptions.AuthenticationRecord = authRecord;
            return new InteractiveBrowserCredential(credentialOptions);
        }
        else
        {
            // create and cache the token
            var browserCredential = new InteractiveBrowserCredential(credentialOptions);
            var tokenRequestContext = new TokenRequestContext(_scopes);
            var authRecord = browserCredential.Authenticate(tokenRequestContext);
            using var authRecordStream = File.OpenWrite(_tokenPath);
            authRecord.Serialize(authRecordStream);
            return browserCredential;
        }
    }

    public void Logout()
    {
        File.Delete(_tokenPath);
        _graphClient = null;
    }
}