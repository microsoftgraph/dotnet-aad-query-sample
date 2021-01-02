using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.Graph;

namespace MsGraph_Samples.Services
{
    public class FakeGraphDataService : IGraphDataService
    {
        public Uri? LastUrl => new("https://graph.microsoft.com/beta/users?$count=true");
        public string? PowerShellCmdLet => "Get-MgUser -ConsistencyLevel eventual -count countVariable";
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

        public Task<IGraphServiceUsersCollectionPage> GetAppOwnersAsUsersAsync(string id, string select, string filter, string orderBy, string search)
        {
            return Task.FromResult((IGraphServiceUsersCollectionPage)new GraphServiceUsersCollectionPage());
        }

        public Task<IGraphServiceGroupsCollectionPage> GetTransitiveMemberOfAsGroupsAsync(string id, string select, string filter, string orderBy, string search)
        {
            return Task.FromResult((IGraphServiceGroupsCollectionPage)new GraphServiceGroupsCollectionPage());
        }

        public Task<IGraphServiceUsersCollectionPage> GetTransitiveMembersAsUsersAsync(string id, string select, string filter, string orderBy, string search)
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
    }
}