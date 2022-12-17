using Microsoft.Graph;
using Microsoft.Graph.Models;

namespace MsGraph_Samples.Services
{
    public class FakeGraphDataService : IGraphDataService
    {
        public string? LastUrl => "https://graph.microsoft.com/beta/users?$count=true";
        private static IList<User> Users => new[]
        {
            new User { Id = "1", DisplayName = "Luca Spolidoro", Mail = "a@b.c" },
            new User { Id = "2", DisplayName = "Pino Quercia", Mail = "pino@quercia.com" },
            new User { Id = "3", DisplayName = "Test Test", Mail = "test@test.com" }
        };

        public Task<User> GetMe(string[] select)
        {
            return Task.FromResult(Users[0]);
        }
        
        public Task<User> GetUser(string id, string[] select)
        {
            throw new NotImplementedException();
        }

        public Task<ApplicationCollectionResponse> GetApplicationsAsync(string[] select, string filter, string[] orderBy, string search)
        {
            return Task.FromResult(new ApplicationCollectionResponse());
        }

        public Task<DeviceCollectionResponse> GetDevicesAsync(string[] select, string filter, string[] orderBy, string search)
        {
            return Task.FromResult(new DeviceCollectionResponse());
        }

        public Task<GroupCollectionResponse> GetGroupsAsync(string[] select, string filter, string[] orderBy, string search)
        {
            return Task.FromResult(new GroupCollectionResponse());
        }

        public Task<UserCollectionResponse> GetAppOwnersAsUsersAsync(string id)
        {
            return Task.FromResult(new UserCollectionResponse());
        }

        public Task<GroupCollectionResponse> GetTransitiveMemberOfAsGroupsAsync(string id)
        {
            return Task.FromResult(new GroupCollectionResponse());
        }

        public Task<UserCollectionResponse> GetTransitiveMembersAsUsersAsync(string id)
        {
            return Task.FromResult(new UserCollectionResponse());
        }

        public Task<UserCollectionResponse> GetUsersAsync(string[] select, string filter, string[] orderBy, string search)
        {
            return Task.FromResult(new UserCollectionResponse());
        }

        public Task<int?> GetUsersRawCountAsync(string filter, string search)
        {
            int? count = Users.Count;
            return Task.FromResult(count);
        }
    }
}