using Microsoft.Graph;

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

        public Task<User> GetMe()
        {
            return Task.FromResult(Users[0]);
        }

        public Task<IGraphServiceApplicationsCollectionPage> GetApplicationsAsync(string select, string filter, string orderBy, string search)
        {
            return Task.FromResult((IGraphServiceApplicationsCollectionPage)new GraphServiceApplicationsCollectionPage());
        }

        public Task<IGraphServiceDevicesCollectionPage> GetDevicesAsync(string select, string filter, string orderBy, string search)
        {
            return Task.FromResult((IGraphServiceDevicesCollectionPage)new GraphServiceDevicesCollectionPage());
        }

        public Task<IGraphServiceGroupsCollectionPage> GetGroupsAsync(string select, string filter, string orderBy, string search)
        {
            return Task.FromResult((IGraphServiceGroupsCollectionPage)new GraphServiceGroupsCollectionPage());
        }

        public Task<IGraphServiceUsersCollectionPage> GetAppOwnersAsUsersAsync(string id)
        {
            return Task.FromResult((IGraphServiceUsersCollectionPage)new GraphServiceUsersCollectionPage());
        }

        public Task<IGraphServiceGroupsCollectionPage> GetTransitiveMemberOfAsGroupsAsync(string id)
        {
            return Task.FromResult((IGraphServiceGroupsCollectionPage)new GraphServiceGroupsCollectionPage());
        }

        public Task<IGraphServiceUsersCollectionPage> GetTransitiveMembersAsUsersAsync(string id)
        {
            return Task.FromResult((IGraphServiceUsersCollectionPage)new GraphServiceUsersCollectionPage());
        }

        public Task<IGraphServiceUsersCollectionPage> GetUsersAsync(string select, string filter, string orderBy, string search)
        {
            return Task.FromResult((IGraphServiceUsersCollectionPage)new GraphServiceUsersCollectionPage());
        }

        public Task<int> GetUsersRawCountAsync(string filter, string search)
        {
            return Task.FromResult(Users.Count);
        }

        public Task<User> GetUser(string id, string select)
        {
            throw new NotImplementedException();
        }
    }
}