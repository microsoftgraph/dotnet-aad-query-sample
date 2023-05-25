using Microsoft.Extensions.Configuration;

namespace MsGraphSamples.Services;

public static class SecretConfig
{
    private static readonly IConfiguration _configuration = new ConfigurationBuilder().AddUserSecrets<AuthService>().Build();

    private const string _helpUrl = "https://docs.microsoft.com/aspnet/core/security/app-secrets?tabs=windows";
    private static readonly FileFormatException _configException = new($"Missing or invalid secrets.json\nMake sure you created one: {_helpUrl}");

    public static string ClientId => _configuration["clientId"] ?? throw _configException;
}