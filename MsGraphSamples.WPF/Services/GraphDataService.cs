// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using System.Collections.Generic;
using System.Net;
using System.Threading.Tasks;
using Microsoft.Graph;

namespace MsGraph_Samples.Services
{
    public interface IGraphDataService
    {
        Task<User> GetMe();
        Task<IGraphServiceApplicationsCollectionPage> GetApplicationsAsync(string filter, string search, string select, string orderBy);
        Task<IGraphServiceDevicesCollectionPage> GetDevicesAsync(string filter, string search, string select, string orderBy);
        Task<IGraphServiceGroupsCollectionPage> GetGroupsAsync(string filter, string search, string select, string orderBy);
        Task<IGraphServiceUsersCollectionPage> GetUsersAsync(string filter, string search, string select, string orderBy);
        Task<IGraphServiceGroupsCollectionPage> GetTransitiveMemberOfAsGroupsAsync(string id);
        Task<IGraphServiceUsersCollectionPage> GetTransitiveMembersAsUsersAsync(string id);
        Task<IGraphServiceUsersCollectionPage> GetAppOwnersAsUsersAsync(string id);
        Task<int> GetUsersRawCountAsync(string filter, string search);

        string? LastUrl { get; }
    }

    public class GraphDataService : IGraphDataService
    {
        // Required for Advanced Queries
        private readonly QueryOption OdataCount = new("$count", "true");

        // Required for Advanced Queries
        private readonly HeaderOption EventualConsistency = new("ConsistencyLevel", "eventual");

        public string? LastUrl { get; private set; } = null;

        private readonly IGraphServiceClient _graphClient;

        public GraphDataService(IGraphServiceClient graphClient)
        {
            _graphClient = graphClient;
        }

        private void AddAdvancedOptions(IBaseRequest request, string filter = "", string search = "", string select = "", string orderBy = "")
        {
            request.QueryOptions.Add(OdataCount);
            request.Headers.Add(EventualConsistency);

            if (!filter.IsNullOrEmpty())
                request.QueryOptions.Add(GetOption("filter", filter));

            if (!orderBy.IsNullOrEmpty())
                request.QueryOptions.Add(GetOption("orderBy", orderBy));

            if (!select.IsNullOrEmpty())
                request.QueryOptions.Add(GetOption("select", select));

            if (!search.IsNullOrEmpty())
                request.QueryOptions.Add(GetOption("search", search));

            LastUrl = WebUtility.UrlDecode(request.GetHttpRequestMessage().RequestUri?.AbsoluteUri);

            static QueryOption GetOption(string name, string value)
            {
                var encodedValue = WebUtility.UrlEncode(value);
                return new QueryOption($"${name}", encodedValue);
            }
        }

        public Task<User> GetMe()
        {
            return _graphClient.Me.Request().GetAsync();
        }

        public Task<IGraphServiceApplicationsCollectionPage> GetApplicationsAsync(string filter, string search, string select, string orderBy)
        {
            var request = _graphClient.Applications.Request();
            AddAdvancedOptions(request, filter, search, select, orderBy);

            return request.GetAsync();
        }

        public Task<IGraphServiceDevicesCollectionPage> GetDevicesAsync(string filter, string search, string select, string orderBy)
        {
            var request = _graphClient.Devices.Request();
            AddAdvancedOptions(request, filter, search, select, orderBy);

            return request.GetAsync();
        }
        public Task<IGraphServiceGroupsCollectionPage> GetGroupsAsync(string filter, string search, string select, string orderBy)
        {
            var request = _graphClient.Groups.Request();
            AddAdvancedOptions(request, filter, search, select, orderBy);

            return request.GetAsync();
        }

        public Task<IGraphServiceUsersCollectionPage> GetUsersAsync(string filter, string search, string select, string orderBy)
        {
            var request = _graphClient.Users.Request();
            AddAdvancedOptions(request, filter, search, select, orderBy);

            return request.GetAsync();
        }

        public Task<IGraphServiceGroupsCollectionPage> GetTransitiveMemberOfAsGroupsAsync(string id)
        {
            var requestUrl = _graphClient.Users[id].TransitiveMemberOf
                .AppendSegmentToRequestUrl("microsoft.graph.group"); // OData Cast
            var request = new GraphServiceGroupsCollectionRequestBuilder(requestUrl, _graphClient).Request();
            AddAdvancedOptions(request);

            return request.GetAsync();
        }

        public Task<IGraphServiceUsersCollectionPage> GetTransitiveMembersAsUsersAsync(string id)
        {
            var requestUrl = _graphClient.Groups[id].TransitiveMembers
                .AppendSegmentToRequestUrl("microsoft.graph.user"); // OData Cast
            var request = new GraphServiceUsersCollectionRequestBuilder(requestUrl, _graphClient).Request();
            AddAdvancedOptions(request);

            return request.GetAsync();
        }

        public Task<IGraphServiceUsersCollectionPage> GetAppOwnersAsUsersAsync(string id)
        {
            var requestUrl = _graphClient.Applications[id].Owners
                .AppendSegmentToRequestUrl("microsoft.graph.user"); // OData Cast
            var request = new GraphServiceUsersCollectionRequestBuilder(requestUrl, _graphClient).Request();
            AddAdvancedOptions(request);

            return request.GetAsync();
        }

        public async Task<int> GetUsersRawCountAsync(string filter, string search)
        {
            var requestUrl = _graphClient.Users.AppendSegmentToRequestUrl("$count");
            var request = new BaseRequest(requestUrl, _graphClient);
            AddAdvancedOptions(request, filter, search);

            var response = await _graphClient.HttpProvider.SendAsync(request.GetHttpRequestMessage());
            var userCount = await response.Content.ReadAsStringAsync();

            return int.Parse(userCount);
        }
    }
}