using System;
using System.Threading.Tasks;
using Microsoft.Graph;
using Microsoft.Identity.Client;

namespace MsGraph_Samples.Services
{
    public class FakeAuthService : IAuthService
    {
        public GraphServiceClient GetServiceClient() => throw new NotImplementedException();

        public Task Logout() => throw new NotImplementedException();
    }
}