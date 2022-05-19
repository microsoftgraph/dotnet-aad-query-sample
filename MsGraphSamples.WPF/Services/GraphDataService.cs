// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using Microsoft.Graph;
using System.Net;

namespace MsGraph_Samples.Services
{
    public interface IGraphDataService
    {
        string? LastUrl { get; }
        Task<User> GetMe(string select);
        Task<IGraphServiceApplicationsCollectionPage> GetApplicationsAsync(string select, string filter, string orderBy, string search);
        Task<IGraphServiceDevicesCollectionPage> GetDevicesAsync(string select, string filter, string orderBy, string search);
        Task<IGraphServiceGroupsCollectionPage> GetGroupsAsync(string select, string filter, string orderBy, string search);
        Task<IGraphServiceUsersCollectionPage> GetUsersAsync(string select, string filter, string orderBy, string search);
        Task<IGraphServiceGroupsCollectionPage> GetTransitiveMemberOfAsGroupsAsync(string id);
        Task<IGraphServiceUsersCollectionPage> GetTransitiveMembersAsUsersAsync(string id);
        Task<IGraphServiceUsersCollectionPage> GetAppOwnersAsUsersAsync(string id);
        Task<int> GetUsersRawCountAsync(string filter, string search);
        Task<User> GetUser(string id, string select);
    }

    public class GraphDataService : IGraphDataService
    {
        public string? LastUrl { get; private set; } = null;

        private readonly GraphServiceClient _graphClient;

        public GraphDataService(GraphServiceClient graphClient)
        {
            _graphClient = graphClient;
        }

        private static IEnumerable<QueryOption> GetAdvancedOptions(string? search = null)
        {
            if (!string.IsNullOrEmpty(search))
                yield return new QueryOption("$search", WebUtility.UrlEncode(search));

            yield return new QueryOption("$count", "true");
        }

        public Task<User> GetMe(string select)
        {
            return _graphClient.Me
                .Request()
                .Select(select)
                .GetAsync();
        }
        
        public Task<User> GetUser(string id, string select)
        {
            return _graphClient.Users[id]
                    .Request()
                    .Select(select)
                    .GetAsync();
        }

        public Task<IGraphServiceApplicationsCollectionPage> GetApplicationsAsync(string select, string filter, string orderBy, string search)
        {
            var request = _graphClient.Applications
                .Request(GetAdvancedOptions(search))
                .Header("ConsistencyLevel", "eventual")
                .Select(select)
                .Filter(filter)
                .OrderBy(orderBy);

            LastUrl = WebUtility.UrlDecode(request.GetHttpRequestMessage().RequestUri?.AbsoluteUri);
            return request.GetAsync();
        }

        public Task<IGraphServiceDevicesCollectionPage> GetDevicesAsync(string select, string filter, string orderBy, string search)
        {
            var request = _graphClient.Devices
                .Request(GetAdvancedOptions(search))
                .Header("ConsistencyLevel", "eventual")
                .Select(select)
                .Filter(filter)
                .OrderBy(orderBy);

            LastUrl = WebUtility.UrlDecode(request.GetHttpRequestMessage().RequestUri?.AbsoluteUri);
            return request.GetAsync();
        }
        
        public Task<IGraphServiceGroupsCollectionPage> GetGroupsAsync(string select, string filter, string orderBy, string search)
        {
            var request = _graphClient.Groups
                .Request(GetAdvancedOptions(search))
                .Header("ConsistencyLevel", "eventual")
                .Select(select)
                .Filter(filter)
                .OrderBy(orderBy);

            LastUrl = WebUtility.UrlDecode(request.GetHttpRequestMessage().RequestUri?.AbsoluteUri);
            return request.GetAsync();
        }

        public Task<IGraphServiceUsersCollectionPage> GetUsersAsync(string select, string filter, string orderBy, string search)
        {
            var request = _graphClient.Users
                .Request(GetAdvancedOptions(search))
                .Header("ConsistencyLevel", "eventual")
                .Select(select)
                .Filter(filter)
                .OrderBy(orderBy);

            LastUrl = WebUtility.UrlDecode(request.GetHttpRequestMessage().RequestUri?.AbsoluteUri);
            return request.GetAsync();
        }

        public Task<IGraphServiceGroupsCollectionPage> GetTransitiveMemberOfAsGroupsAsync(string id)
        {
            var requestUrl = _graphClient.Users[id].TransitiveMemberOf
                .AppendSegmentToRequestUrl("microsoft.graph.group"); // OData Cast

            var request = new GraphServiceGroupsCollectionRequestBuilder(requestUrl, _graphClient)
                .Request(GetAdvancedOptions(null))
                .Header("ConsistencyLevel", "eventual");

            LastUrl = WebUtility.UrlDecode(request.GetHttpRequestMessage().RequestUri?.AbsoluteUri);
            return request.GetAsync();
        }

        public Task<IGraphServiceUsersCollectionPage> GetTransitiveMembersAsUsersAsync(string id)
        {
            var requestUrl = _graphClient.Groups[id].TransitiveMembers
                .AppendSegmentToRequestUrl("microsoft.graph.user"); // OData Cast

            var request = new GraphServiceUsersCollectionRequestBuilder(requestUrl, _graphClient)
                .Request(GetAdvancedOptions())
                .Header("ConsistencyLevel", "eventual");

            LastUrl = WebUtility.UrlDecode(request.GetHttpRequestMessage().RequestUri?.AbsoluteUri);
            return request.GetAsync();
        }

        public Task<IGraphServiceUsersCollectionPage> GetAppOwnersAsUsersAsync(string id)
        {
            var requestUrl = _graphClient.Applications[id].Owners
                .AppendSegmentToRequestUrl("microsoft.graph.user"); // OData Cast
            var request = new GraphServiceUsersCollectionRequestBuilder(requestUrl, _graphClient)
                .Request(GetAdvancedOptions())
                .Header("ConsistencyLevel", "eventual");

            LastUrl = WebUtility.UrlDecode(request.GetHttpRequestMessage().RequestUri?.AbsoluteUri);
            return request.GetAsync();
        }

        public async Task<int> GetUsersRawCountAsync(string filter, string search)
        {
            var queryOptions = new[]
            {
                new QueryOption("$filter", WebUtility.UrlEncode(filter)),
                new QueryOption("$search", WebUtility.UrlEncode(search))
            };

            var requestUrl = _graphClient.Users.AppendSegmentToRequestUrl("$count");
            var request = new BaseRequest(requestUrl, _graphClient, queryOptions)
                .Header("ConsistencyLevel", "eventual");

            LastUrl = WebUtility.UrlDecode(request.GetHttpRequestMessage().RequestUri?.AbsoluteUri);

            var response = await _graphClient.HttpProvider.SendAsync(request.GetHttpRequestMessage());
            var userCount = await response.Content.ReadAsStringAsync();
            return int.Parse(userCount);
        }
    }
}