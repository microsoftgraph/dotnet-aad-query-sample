using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Microsoft.Graph;

namespace MsGraph_Samples.Services
{
    public class FakeGraphDataService : IGraphDataService
    {
        public Uri? LastCall => throw new NotImplementedException();

        public Task<IEnumerable<DirectoryObject>?> GetApplicationsAsync(string filter, string search, string select, string orderBy)
        {
            return Task.FromResult<IEnumerable<DirectoryObject>?>(null);
        }

        public Task<IEnumerable<DirectoryObject>?> GetDevicesAsync(string filter, string search, string select, string orderBy)
        {
            return Task.FromResult<IEnumerable<DirectoryObject>?>(null);
        }

        public Task<IEnumerable<DirectoryObject>?> GetGroupsAsync(string filter, string search, string select, string orderBy)
        {
            return Task.FromResult<IEnumerable<DirectoryObject>?>(null);
        }

        public Task<IEnumerable<DirectoryObject>?> GetAppOwnersAsUsersAsync(string id)
        {
            return Task.FromResult<IEnumerable<DirectoryObject>?>(null);
        }

        public Task<IEnumerable<DirectoryObject>?> GetTransitiveMemberOfAsGroupsAsync(string id)
        {
            return Task.FromResult<IEnumerable<DirectoryObject>?>(null);
        }

        public Task<IEnumerable<DirectoryObject>?> GetTransitiveMembersAsUsersAsync(string id)
        {
            return Task.FromResult<IEnumerable<DirectoryObject>?>(null);
        }

        public Task<IEnumerable<DirectoryObject>?> GetUsersAsync(string filter, string search, string select, string orderBy)
        {
            var directoryObjects = new[]
            {
                new DirectoryObject { Id = "1" },
                new DirectoryObject { Id = "2" },
                new DirectoryObject { Id = "3" }
            };

            return Task.FromResult<IEnumerable<DirectoryObject>?>(directoryObjects);
        }

        public Task<long> GetUsersRawCountAsync(string filter, string search)
        {
            return Task.FromResult(0L);
        }
    }
}