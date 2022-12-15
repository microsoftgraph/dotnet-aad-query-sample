using Microsoft.Graph;

namespace MsGraph_Samples.Services
{
    public class FakeAuthService : IAuthService
    {
        public GraphServiceClient GraphClient => throw new NotImplementedException();

        void IAuthService.Logout()
        {
            throw new NotImplementedException();
        }
    }
}