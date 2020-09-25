// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using System;
using System.Collections.Generic;
using System.Text.Encodings.Web;
using System.Threading.Tasks;
using Microsoft.Graph;

namespace MsGraph_Samples.Services
{
    public interface IGraphDataService
    {
        Task<IEnumerable<DirectoryObject>?> GetApplicationsAsync(string filter, string search, string select, string orderBy);
        Task<IEnumerable<DirectoryObject>?> GetDevicesAsync(string filter, string search, string select, string orderBy);
        Task<IEnumerable<DirectoryObject>?> GetGroupsAsync(string filter, string search, string select, string orderBy);
        Task<IEnumerable<DirectoryObject>?> GetUsersAsync(string filter, string search, string select, string orderBy);
        Task<IEnumerable<DirectoryObject>?> GetTransitiveMemberOfAsGroupsAsync(string id);
        Task<IEnumerable<DirectoryObject>?> GetTransitiveMembersAsUsersAsync(string id);
        Task<IEnumerable<DirectoryObject>?> GetAppOwnersAsUsersAsync(string id);
        
        Task<long> GetUsersRawCountAsync(string filter, string search);
        Uri? LastCall { get; }
    }

    public class GraphDataService : IGraphDataService
    {
        // Required for Advanced Queries
        private readonly QueryOption OdataCount = new QueryOption("$count", "true");
        // Required for Advanced Queries
        private readonly Option[] EventualConsistency = new[] { new HeaderOption("ConsistencyLevel", "eventual") };

        /// <summary>
        /// Used for to show the full URL in case of errors
        /// </summary>
        public Uri? LastCall { get; private set; } = null;

        private readonly IGraphServiceClient _graphClient;

        public GraphDataService(IGraphServiceClient graphClient)
        {
            _graphClient = graphClient;
        }

        private void AddAdvancedOptions(IBaseRequest request, string? filter = null, string? search = null, string? select = null, string? orderBy = null)
        {
            request.QueryOptions.Add(OdataCount);

            if (!string.IsNullOrEmpty(filter))
                request.QueryOptions.Add(GetOption("filter", filter));

            if (!string.IsNullOrEmpty(orderBy))
                request.QueryOptions.Add(GetOption("orderBy", orderBy));

            if (!string.IsNullOrEmpty(select))
                request.QueryOptions.Add(GetOption("select", select));

            if (!string.IsNullOrEmpty(search))
                request.QueryOptions.Add(GetSearchOption(search));

            LastCall = request.GetHttpRequestMessage().RequestUri;

            static QueryOption GetOption(string name, string value) => new QueryOption($"${name}", $"{UrlEncoder.Default.Encode(value)}");
            static QueryOption GetSearchOption(string value) => new QueryOption("$search", $"\"{UrlEncoder.Default.Encode(value)}\"");
        }

        public async Task<IEnumerable<DirectoryObject>?> GetDevicesAsync(string filter, string search, string select, string orderBy)
        {
            var request = _graphClient.Devices.Request(EventualConsistency);
            AddAdvancedOptions(request, filter, search, select, orderBy);

            return await request.GetAsync();
        }

        public async Task<IEnumerable<DirectoryObject>?> GetUsersAsync(string filter, string search, string select, string orderBy)
        {
            var request = _graphClient.Users.Request(EventualConsistency);
            AddAdvancedOptions(request, filter, search, select, orderBy);

            return await request.GetAsync();
        }

        public async Task<IEnumerable<DirectoryObject>?> GetGroupsAsync(string filter, string search, string select, string orderBy)
        {
            var request = _graphClient.Groups.Request(EventualConsistency);
            AddAdvancedOptions(request, filter, search, select, orderBy);

            return await request.GetAsync();
        }

        public async Task<IEnumerable<DirectoryObject>?> GetApplicationsAsync(string filter, string search, string select, string orderBy)
        {
            var request = _graphClient.Applications.Request(EventualConsistency);
            AddAdvancedOptions(request, filter, search, select, orderBy);

            return await request.GetAsync();
        }

        public async Task<IEnumerable<DirectoryObject>?> GetTransitiveMemberOfAsGroupsAsync(string id)
        {
            var requestUrl = _graphClient.Users[id].TransitiveMemberOf
                .AppendSegmentToRequestUrl("microsoft.graph.group"); // OData Cast
            var request = new GraphServiceGroupsCollectionRequest(requestUrl, _graphClient, EventualConsistency);
            AddAdvancedOptions(request);

            return await request.GetAsync();
        }

        public async Task<IEnumerable<DirectoryObject>?> GetTransitiveMembersAsUsersAsync(string id)
        {
            var requestUrl = _graphClient.Groups[id].TransitiveMembers
                .AppendSegmentToRequestUrl("microsoft.graph.user"); // OData Cast
            var request = new GraphServiceUsersCollectionRequest(requestUrl, _graphClient, EventualConsistency);
            AddAdvancedOptions(request);

            return await request.GetAsync();
        }

        public async Task<IEnumerable<DirectoryObject>?> GetAppOwnersAsUsersAsync(string id)
        {
            var requestUrl = _graphClient.Applications[id].Owners
                .AppendSegmentToRequestUrl("microsoft.graph.user"); // OData Cast
            var request = new GraphServiceUsersCollectionRequest(requestUrl, _graphClient, EventualConsistency);
            AddAdvancedOptions(request);

            return await request.GetAsync();
        }

        public async Task<long> GetUsersRawCountAsync(string filter, string search)
        {
            var requestUrl = _graphClient.Users.AppendSegmentToRequestUrl("$count");
            var request = new GraphServiceUsersCollectionRequest(requestUrl, _graphClient, EventualConsistency);
            AddAdvancedOptions(request, filter, search);

            var requestMessage = request.GetHttpRequestMessage();
            var responseMessage = await _graphClient.HttpProvider.SendAsync(requestMessage);
            
            var userCount = await responseMessage.Content.ReadAsStringAsync();
            return long.Parse(userCount);
        }
    }
}