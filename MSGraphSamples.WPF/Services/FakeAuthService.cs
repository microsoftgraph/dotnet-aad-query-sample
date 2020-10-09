using System;
using System.Net.Http;
using System.Threading.Tasks;
using Microsoft.Graph;
using Microsoft.Identity.Client;

namespace MsGraph_Samples.Services
{
    public class FakeAuthService : IAuthService
    {
        public IAccount? Account => null;

        public event Action? AuthenticationSuccessful;

        public Task AuthenticateRequestAsync(HttpRequestMessage request)
        {
            AuthenticationSuccessful?.Invoke();
            return Task.CompletedTask;
        }

        public GraphServiceClient GetServiceClient()
        {
            throw new NotImplementedException();
        }

        public Task Logout()
        {
            throw new NotImplementedException();
        }
    }
}
