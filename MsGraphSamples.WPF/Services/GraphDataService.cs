// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.IdentityModel.Tokens;
using System.Net;

namespace MsGraph_Samples.Services
{
    public interface IGraphDataService
    {
        string? LastUrl { get; }
        Task<User> GetMe(string[] select);
        Task<ApplicationCollectionResponse> GetApplicationsAsync(string[] select, string filter, string[] orderBy, string search);
        Task<DeviceCollectionResponse> GetDevicesAsync(string[] select, string filter, string[] orderBy, string search);
        Task<GroupCollectionResponse> GetGroupsAsync(string[] select, string filter, string[] orderBy, string search);
        Task<UserCollectionResponse> GetUsersAsync(string[] select, string filter, string[] orderBy, string search);
        Task<GroupCollectionResponse> GetTransitiveMemberOfAsGroupsAsync(string id);
        Task<UserCollectionResponse> GetTransitiveMembersAsUsersAsync(string id);
        Task<UserCollectionResponse> GetAppOwnersAsUsersAsync(string id);
        Task<int?> GetUsersRawCountAsync(string filter, string search);
        Task<User> GetUser(string id, string[] select);
    }

    public class GraphDataService : IGraphDataService
    {
        public string? LastUrl { get; private set; } = null;

        private readonly GraphServiceClient _graphClient;

        private readonly Dictionary<string, string> EventualConsistencyHeader = new()
        {
            { "ConsistencyLevel", "eventual" }
        };

        public GraphDataService(GraphServiceClient graphClient)
        {
            _graphClient = graphClient;
        }

        public Task<User> GetMe(string[] select)
        {
            return _graphClient.Me
                .GetAsync(rc => rc.QueryParameters.Select = select);
        }

        public Task<User> GetUser(string id, string[] select)
        {
            return _graphClient.Users[id]
                .GetAsync(rc => rc.QueryParameters.Select = select);
        }

        //public IAsyncEnumerable<Application> GetApplicationsAsync(string[] select, string filter, string[] orderBy, string search)
        //{
        //    var requestInfo = _graphClient.Applications
        //        .CreateGetRequestInformation(rc =>
        //        {
        //            rc.Headers = EventualConsistencyHeader;
        //            rc.QueryParameters.Search = WebUtility.UrlEncode(search);
        //            rc.QueryParameters.Select = select;
        //            rc.QueryParameters.Filter = filter;
        //            rc.QueryParameters.Orderby = orderBy;
        //        });

        //    LastUrl = requestInfo.URI.AbsoluteUri;
        //    return requestInfo.ToAsyncEnumerable<Application>(_graphClient.RequestAdapter);
        //}

        public Task<ApplicationCollectionResponse> GetApplicationsAsync(string[] select, string filter, string[] orderBy, string search)
        {
            return _graphClient.Applications
                .GetAsync(rc =>
                {
                    rc.Headers = EventualConsistencyHeader;
                    rc.QueryParameters.Count = true;
                    if (!search.IsNullOrEmpty())
                        rc.QueryParameters.Search = WebUtility.UrlEncode(search);
                    rc.QueryParameters.Select = select;
                    rc.QueryParameters.Filter = filter;
                    rc.QueryParameters.Orderby = orderBy;
                });

            //LastUrl = WebUtility.UrlDecode(request.GetHttpRequestMessage().RequestUri?.AbsoluteUri);
            //return request.GetAsync();
        }

        public Task<DeviceCollectionResponse> GetDevicesAsync(string[] select, string filter, string[] orderBy, string search)
        {
            return _graphClient.Devices
                .GetAsync(rc =>
                {
                    rc.Headers = EventualConsistencyHeader;
                    rc.QueryParameters.Count = true;
                    if (!search.IsNullOrEmpty())
                        rc.QueryParameters.Search = WebUtility.UrlEncode(search);
                    rc.QueryParameters.Select = select;
                    rc.QueryParameters.Filter = filter;
                    rc.QueryParameters.Orderby = orderBy;
                });

            //LastUrl = WebUtility.UrlDecode(request.GetHttpRequestMessage().RequestUri?.AbsoluteUri);
            //return request.GetAsync();
        }

        public Task<GroupCollectionResponse> GetGroupsAsync(string[] select, string filter, string[] orderBy, string search)
        {
            return _graphClient.Groups
                .GetAsync(rc =>
                {
                    rc.Headers = EventualConsistencyHeader;
                    rc.QueryParameters.Count = true;
                    if (!search.IsNullOrEmpty())
                        rc.QueryParameters.Search = WebUtility.UrlEncode(search);
                    rc.QueryParameters.Select = select;
                    rc.QueryParameters.Filter = filter;
                    rc.QueryParameters.Orderby = orderBy;
                });

            //LastUrl = WebUtility.UrlDecode(request.GetHttpRequestMessage().RequestUri?.AbsoluteUri);
            //return request.GetAsync();
        }

        public  Task<UserCollectionResponse> GetUsersAsync(string[] select, string filter, string[] orderBy, string search)
        {
            var requestInfo = _graphClient.Users
                .CreateGetRequestInformation(rc =>
                {
                    rc.Headers = EventualConsistencyHeader;
                    rc.QueryParameters.Count = true;
                    if(!search.IsNullOrEmpty())
                        rc.QueryParameters.Search = WebUtility.UrlEncode(search);
                    rc.QueryParameters.Select = select;
                    rc.QueryParameters.Filter = filter;
                    rc.QueryParameters.Orderby = orderBy;
                });

            //LastUrl = WebUtility.UrlDecode(requestInfo.URI.AbsoluteUri);
            return _graphClient.RequestAdapter.SendAsync(requestInfo, (parseNode) => new UserCollectionResponse());

            //return request.GetAsync();
        }

        public Task<GroupCollectionResponse> GetTransitiveMemberOfAsGroupsAsync(string id)
        {
            return _graphClient.Users[id]
                .TransitiveMemberOf.Group
                .GetAsync(rc =>
                {
                    rc.Headers = EventualConsistencyHeader;
                    rc.QueryParameters.Count = true;
                });

            //LastUrl = WebUtility.UrlDecode(request.GetHttpRequestMessage().RequestUri?.AbsoluteUri);
            //return request.GetAsync();
        }

        public Task<UserCollectionResponse> GetTransitiveMembersAsUsersAsync(string id)
        {
            return _graphClient.Groups[id]
                .TransitiveMembers.User
                .GetAsync(rc =>
                {
                    rc.Headers = EventualConsistencyHeader;
                    rc.QueryParameters.Count = true;
                });

            //LastUrl = WebUtility.UrlDecode(request.GetHttpRequestMessage().RequestUri?.AbsoluteUri);
            //return request.GetAsync();
        }

        public Task<UserCollectionResponse> GetAppOwnersAsUsersAsync(string id)
        {
            return _graphClient.Applications[id]
                .Owners.User
                .GetAsync(rc =>
                {
                    rc.Headers = EventualConsistencyHeader;
                    rc.QueryParameters.Count = true;
                });

            //LastUrl = WebUtility.UrlDecode(request.GetHttpRequestMessage().RequestUri?.AbsoluteUri);
            //return request.GetAsync();
        }

        public Task<int?> GetUsersRawCountAsync(string filter, string search)
        {
            return _graphClient.Users.Count
               .GetAsync(rc =>
               {
                   rc.Headers = EventualConsistencyHeader;
                   //rc.QueryParameters.Search = WebUtility.UrlEncode(search);
                   //rc.QueryParameters.Filter = filter;
               });

            //LastUrl = WebUtility.UrlDecode(request.GetHttpRequestMessage().RequestUri?.AbsoluteUri);
            //return request.GetAsync();
        }
    }
}