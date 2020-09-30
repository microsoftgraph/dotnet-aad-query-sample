using System.IO;
using Microsoft.Extensions.Configuration;

namespace MsGraph_Samples.Helpers
{
    public static class SecretConfig
    {
        private readonly static IConfiguration _configuration = new ConfigurationBuilder().AddUserSecrets<App>().Build();

        private const string _helpUrl = "https://docs.microsoft.com/aspnet/core/security/app-secrets?view=aspnetcore-3.1&tabs=windows";
        private readonly static FileFormatException _configException = new FileFormatException($"Missing or invalid secrets.json\nMake sure you created one: {_helpUrl}");

        public static string ClientId => _configuration["clientId"] ?? throw _configException;
    }
}