using System.IO;
using Microsoft.Extensions.Configuration;
using MsGraph_Samples.Services;

namespace MsGraph_Samples.Helpers
{
    public static class SecretConfig
    {
        private readonly static IConfiguration _configuration = new ConfigurationBuilder().AddUserSecrets<GraphDataService>().Build();

        private const string _helpUrl = "https://docs.microsoft.com/aspnet/core/security/app-secrets?tabs=windows";
        private readonly static FileFormatException _configException = new($"Missing or invalid secrets.json\nMake sure you created one: {_helpUrl}");

        public static string ClientId => _configuration["clientId"] ?? throw _configException;
    }
}