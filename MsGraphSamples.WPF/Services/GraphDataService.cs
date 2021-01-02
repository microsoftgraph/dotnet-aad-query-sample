// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using System;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Graph;

namespace MsGraph_Samples.Services
{
    public interface IGraphDataService
    {
        Task<User> GetMe();
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
                    cmdLet.Append(param switch
                    {
                        "$count" => "-CountVariable countVar ",
                        "$select" => $"-Property {queryString[param]} ",
                        "$filter" => $"-Filter {queryString[param]} ",
                        "$orderBy" => $"-Sort {queryString[param]} ",
                        "$search" => $"-Search '{queryString[param]}' ",
                        _ => throw new InvalidOperationException("Can't parse QueryString parameter in Url")
                    });
                }

                return cmdLet.ToString().TrimEnd();
            }
        }


        private readonly IGraphServiceClient _graphClient;

        public GraphDataService(IGraphServiceClient graphClient)
        {
            _graphClient = graphClient;
        }

        private void AddAdvancedOptions(IBaseRequest request, string select = "", string filter = "", string orderBy = "", string search = "")
        {
            request.QueryOptions.Add(OdataCount);
            request.Headers.Add(EventualConsistency);

            if (!select.IsNullOrEmpty())
                request.QueryOptions.Add(GetOption("select", select));

            if (!filter.IsNullOrEmpty())
                request.QueryOptions.Add(GetOption("filter", filter));

            if (!orderBy.IsNullOrEmpty())
                request.QueryOptions.Add(GetOption("orderBy", orderBy));

            if (!search.IsNullOrEmpty())
                request.QueryOptions.Add(GetOption("search", search));

            LastUrl = request.GetHttpRequestMessage().RequestUri;

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

        public Task<IGraphServiceApplicationsCollectionPage> GetApplicationsAsync(string select, string filter, string orderBy, string search)
        {
            var request = _graphClient.Applications.Request();
            AddAdvancedOptions(request, select, filter, orderBy, search);

            return request.GetAsync();
        }

        public Task<IGraphServiceDevicesCollectionPage> GetDevicesAsync(string select, string filter, string orderBy, string search)
        {
            var request = _graphClient.Devices.Request();
            AddAdvancedOptions(request, select, filter, orderBy, search);

            return request.GetAsync();
        }
        public Task<IGraphServiceGroupsCollectionPage> GetGroupsAsync(string select, string filter, string orderBy, string search)
        {
            var request = _graphClient.Groups.Request();
            AddAdvancedOptions(request, select, filter, orderBy, search);

            return request.GetAsync();
        }

        public Task<IGraphServiceUsersCollectionPage> GetUsersAsync(string select, string filter, string orderBy, string search)
        {
            var request = _graphClient.Users.Request();
            AddAdvancedOptions(request, select, filter, orderBy, search);

            return request.GetAsync();
        }

        public Task<IGraphServiceGroupsCollectionPage> GetTransitiveMemberOfAsGroupsAsync(string id, string select, string filter, string orderBy, string search)
        {
            var requestUrl = _graphClient.Users[id].TransitiveMemberOf
                .AppendSegmentToRequestUrl("microsoft.graph.group"); // OData Cast
            var request = new GraphServiceGroupsCollectionRequestBuilder(requestUrl, _graphClient).Request();
            AddAdvancedOptions(request, select, filter, orderBy, search);

            return request.GetAsync();
        }

        public Task<IGraphServiceUsersCollectionPage> GetTransitiveMembersAsUsersAsync(string id, string select, string filter, string orderBy, string search)
        {
            var requestUrl = _graphClient.Groups[id].TransitiveMembers
                .AppendSegmentToRequestUrl("microsoft.graph.user"); // OData Cast
            var request = new GraphServiceUsersCollectionRequestBuilder(requestUrl, _graphClient).Request();
            AddAdvancedOptions(request, select, filter, orderBy, search);

            return request.GetAsync();
        }

        public Task<IGraphServiceUsersCollectionPage> GetAppOwnersAsUsersAsync(string id, string select, string filter, string orderBy, string search)
        {
            var requestUrl = _graphClient.Applications[id].Owners
                .AppendSegmentToRequestUrl("microsoft.graph.user"); // OData Cast
            var request = new GraphServiceUsersCollectionRequestBuilder(requestUrl, _graphClient).Request();
            AddAdvancedOptions(request, select, filter, orderBy, search);

            return request.GetAsync();
        }

        public async Task<int> GetUsersRawCountAsync(string filter, string search)
        {
            var requestUrl = _graphClient.Users.AppendSegmentToRequestUrl("$count");
            var request = new BaseRequest(requestUrl, _graphClient);
            AddAdvancedOptions(request, filter: filter, search: search);

            var response = await _graphClient.HttpProvider.SendAsync(request.GetHttpRequestMessage());
            var userCount = await response.Content.ReadAsStringAsync();

            return int.Parse(userCount);
        }
    }
}