// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using Azure.Core;
using Azure.Identity;
using Microsoft.Graph;

namespace MsGraphSamples.Services;

public interface IAuthService
{
    GraphServiceClient GraphClient { get; }
    void Logout();
}

public class AuthService : IAuthService
{
    private const string _tokenPath = "authToken.bin";
    private static readonly string[] _scopes = { "Directory.ReadWrite.All" };

    private GraphServiceClient? _graphClient;
    public GraphServiceClient GraphClient => _graphClient ??= new GraphServiceClient(GetBrowserCredential());

    public void Logout()
    {
        File.Delete(_tokenPath);
        _graphClient = null;
    }

    private static InteractiveBrowserCredential GetBrowserCredential()
    {
        var credentialOptions = new InteractiveBrowserCredentialOptions
        {
            ClientId = SecretConfig.ClientId,
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
}