// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using Azure.Core;
using Azure.Identity;
using Microsoft.Graph;
using MsGraph_Samples.Helpers;

namespace MsGraph_Samples.Services;

public interface IAuthService
{
    GraphServiceClient GraphClient { get; }
    void Logout();
}

public class AuthService : IAuthService
{
    private const string _tokenPath = "authToken.bin";
    public static readonly string[] Scopes = { "Directory.Read.All" };

    private GraphServiceClient? _graphClient;
    public GraphServiceClient GraphClient => _graphClient ??= new GraphServiceClient(GetBrowserCredential());

    public void Logout()
    {
        System.IO.File.Delete(_tokenPath);
        _graphClient = null;
    }

    private static InteractiveBrowserCredential GetBrowserCredential()
    {
        var credentialOptions = new InteractiveBrowserCredentialOptions
        {
            ClientId = SecretConfig.ClientId,
            TokenCachePersistenceOptions = new TokenCachePersistenceOptions() { UnsafeAllowUnencryptedStorage = true }
        };

        if (System.IO.File.Exists(_tokenPath))
        {
            // use the cached token
            using var authRecordStream = System.IO.File.OpenRead(_tokenPath);
            var authRecord = AuthenticationRecord.Deserialize(authRecordStream);
            credentialOptions.AuthenticationRecord = authRecord;
            return new InteractiveBrowserCredential(credentialOptions);
        }
        else
        {
            // create and cache the token
            var browserCredential = new InteractiveBrowserCredential(credentialOptions);
            var tokenRequestContext = new TokenRequestContext(Scopes);
            var authRecord = browserCredential.Authenticate(tokenRequestContext);
            using var authRecordStream = System.IO.File.OpenWrite(_tokenPath);
            authRecord.Serialize(authRecordStream);
            return browserCredential;
        }
    }
}