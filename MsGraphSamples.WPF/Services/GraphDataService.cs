// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using Microsoft.Graph;
using Microsoft.IdentityModel.Tokens;
using System.Net;
using System.Text;

namespace MsGraph_Samples.Services
{
    public interface IGraphDataService
    {
        Task<User> GetMe(string select);
        Task<User> GetUser(string id, string select);
        Task<IGraphServiceApplicationsCollectionPage> GetApplicationsAsync(string select, string filter, string orderBy, string search);
        Task<IGraphServiceDevicesCollectionPage> GetDevicesAsync(string select, string filter, string orderBy, string search);
        Task<IGraphServiceGroupsCollectionPage> GetGroupsAsync(string select, string filter, string orderBy, string search);
        Task<IGraphServiceUsersCollectionPage> GetUsersAsync(string select, string filter, string orderBy, string search);
        Task<IGraphServiceGroupsCollectionPage> GetTransitiveMemberOfAsGroupsAsync(string id, string select, string filter, string orderBy, string search);
        Task<IGraphServiceUsersCollectionPage> GetTransitiveMembersAsUsersAsync(string id, string select, string filter, string orderBy, string search);
        Task<IGraphServiceUsersCollectionPage> GetAppOwnersAsUsersAsync(string id, string select, string filter, string orderBy, string search);
        Task<int> GetUsersRawCountAsync(string filter, string search);

        Uri? LastUrl { get; }
        string? PowerShellCmdLet { get; }
    }

    public class GraphDataService : IGraphDataService
    {
        // Required for Advanced Queries
        private readonly QueryOption OdataCount = new("$count", "true");

        // Required for Advanced Queries
        private readonly HeaderOption EventualConsistency = new("ConsistencyLevel", "eventual");

        public Uri? LastUrl { get; private set; } = null;

        private readonly GraphServiceClient _graphClient;
        public GraphDataService(GraphServiceClient graphClient)
        {
            _graphClient = graphClient;
        }
        public string? PowerShellCmdLet
        {
            get
            {
                if (LastUrl == null)
                    return null;

                // https://github.com/microsoftgraph/msgraph-sdk-powershell/tree/dev/src
                StringBuilder cmdLet = new();

                var segments = LastUrl.AbsolutePath.Split('/').Skip(2).ToArray();

                var isLink = false;
                foreach (var segment in segments)
                {
                    switch (segment)
                    {
                        case "users":
                            cmdLet.Append("Get-MgUser");
                            break;
                        case "groups":
                            cmdLet.Append("Get-MgGroup");
                            break;
                        case "applications":
                            cmdLet.Append("Get-MgApplication");
                            break;
                        case "devices":
                            cmdLet.Append("Get-MgDevice");
                            break;
                        case "transitiveMembers":
                            cmdLet.Append("TransitiveMember");
                            isLink = true;
                            break;
                        case "transitiveMemberOf":
                            cmdLet.Append("TransitiveMemberOf"); // TODO: not every command is like this
                            isLink = true;
                            break;
                        case "owner":
                            cmdLet.Append("Owner");
                            isLink = true;
                            break;
                        //case "$count": // TODO: Raw count not supported?
                        //    cmdLet.Append(" -Count");
                        //    break;
                        default: // id parsed in next switch
                            break;
                    }
                }

                if (isLink) //TODO: ODATA Cast?
                {
                    switch (segments[0])
                    {
                        case "users":
                            cmdLet.Append($" -UserId {segments[1]}");
                            break;
                        case "groups":
                            cmdLet.Append($" -GroupId {segments[1]}");
                            break;
                        case "applications":
                            cmdLet.Append($" -ApplicationId {segments[1]}");
                            break;
                        case "devices":
                            cmdLet.Append($" -DeviceId {segments[1]}");
                            break;
                        default:
                            throw new InvalidOperationException("Can't parse Entity in Url");
                    }
                }
                else //TODO: ConsistencyLevel parameter not supported for links
                {
                    cmdLet.Append(" -ConsistencyLevel eventual");
                }

                cmdLet.Append(' ');

                var queryString = LastUrl.ParseQueryString();
                foreach (string param in queryString)
                {
                    if (queryString[param].IsNullOrEmpty())
                        continue;

                    cmdLet.Append(param switch
                    {
                        "$count" => "-CountVariable countVar ",
                        "$select" => $"-Property {queryString[param]} ",
                        "$filter" => $"-Filter {queryString[param]} ",
                        "$orderby" => $"-Sort {queryString[param]} ",
                        "$search" => $"-Search '{queryString[param]}' ",
                        _ => throw new InvalidOperationException("Can't parse QueryString parameter in Url")
                    });
                }

                return cmdLet.ToString().TrimEnd();
            }
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

            LastUrl = request.GetHttpRequestMessage().RequestUri;
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

            LastUrl = request.GetHttpRequestMessage().RequestUri;
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

            LastUrl = request.GetHttpRequestMessage().RequestUri;
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

            LastUrl = request.GetHttpRequestMessage().RequestUri;
            return request.GetAsync();
        }



        public Task<IGraphServiceGroupsCollectionPage> GetTransitiveMemberOfAsGroupsAsync(string id, string select, string filter, string orderBy, string search)
        {
            var requestUrl = _graphClient.Users[id].TransitiveMemberOf
                .AppendSegmentToRequestUrl("microsoft.graph.group"); // OData Cast

            var request = new GraphServiceGroupsCollectionRequestBuilder(requestUrl, _graphClient)
                .Request(GetAdvancedOptions(null))
                .Header("ConsistencyLevel", "eventual");

            LastUrl = request.GetHttpRequestMessage().RequestUri;
            return request.GetAsync();
        }

        public Task<IGraphServiceUsersCollectionPage> GetTransitiveMembersAsUsersAsync(string id, string select, string filter, string orderBy, string search)
        {
            var requestUrl = _graphClient.Groups[id].TransitiveMembers
                .AppendSegmentToRequestUrl("microsoft.graph.user"); // OData Cast

            var request = new GraphServiceUsersCollectionRequestBuilder(requestUrl, _graphClient)
                .Request(GetAdvancedOptions())
                .Header("ConsistencyLevel", "eventual");

            LastUrl = request.GetHttpRequestMessage().RequestUri;
            return request.GetAsync();
        }

        public Task<IGraphServiceUsersCollectionPage> GetAppOwnersAsUsersAsync(string id, string select, string filter, string orderBy, string search)
        {
            var requestUrl = _graphClient.Applications[id].Owners
                .AppendSegmentToRequestUrl("microsoft.graph.user"); // OData Cast
            var request = new GraphServiceUsersCollectionRequestBuilder(requestUrl, _graphClient)
                .Request(GetAdvancedOptions())
                .Header("ConsistencyLevel", "eventual");

            LastUrl = request.GetHttpRequestMessage().RequestUri;
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

            LastUrl = request.GetHttpRequestMessage().RequestUri;

            var response = await _graphClient.HttpProvider.SendAsync(request.GetHttpRequestMessage());
            var userCount = await response.Content.ReadAsStringAsync();
            return int.Parse(userCount);
        }
    }
}