// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.IdentityModel.Tokens;
using Microsoft.Kiota.Abstractions;
using MsGraph_Samples.Helpers;
using System.Net;

namespace MsGraph_Samples.Services
{
    public interface IGraphDataService
    {
        string? LastUrl { get; }
        Task<User> GetUserAsync(string[] select, string? id = null);
        Task<int?> GetUsersRawCountAsync(string filter, string search);

        IAsyncEnumerable<Application> GetApplications(string[] select, string filter, string[] orderBy, string search);
        Task<ApplicationCollectionResponse> GetApplicationCollectionAsync(string[] select, string filter, string[] orderBy, string search);

        IAsyncEnumerable<Device> GetDevices(string[] select, string filter, string[] orderBy, string search);
        Task<DeviceCollectionResponse> GetDeviceCollectionAsync(string[] select, string filter, string[] orderBy, string search);

        IAsyncEnumerable<Group> GetGroups(string[] select, string filter, string[] orderBy, string search);
        Task<GroupCollectionResponse> GetGroupCollectionAsync(string[] select, string filter, string[] orderBy, string search);

        IAsyncEnumerable<User> GetUsers(string[] select, string filter, string[] orderBy, string search);
        Task<UserCollectionResponse> GetUserCollectionAsync(string[] select, string filter, string[] orderBy, string search);

        IAsyncEnumerable<Group> GetTransitiveMemberOfAsGroups(string id);
        Task<GroupCollectionResponse> GetTransitiveMemberOfAsGroupCollectionAsync(string id);

        IAsyncEnumerable<User> GetTransitiveMembersAsUsers(string id);
        Task<UserCollectionResponse> GetTransitiveMembersAsUserCollectionAsync(string id);

        IAsyncEnumerable<User> GetAppOwnersAsUsers(string id);
        Task<UserCollectionResponse> GetAppOwnersAsUserCollectionAsync(string id);
    }

    public class GraphDataService : IGraphDataService
    {
        public string? LastUrl { get; private set; } = null;

        private readonly GraphServiceClient _graphClient;

        private readonly RequestHeaders EventualConsistencyHeader = new()
        {
            { "ConsistencyLevel", "eventual" }
        };

        public GraphDataService(GraphServiceClient graphClient)
        {
            _graphClient = graphClient;
        }

        public Task<User> GetUserAsync(string[] select, string? id = null)
        {
            return id == null
                ? _graphClient.Me.GetAsync(rc => rc.QueryParameters.Select = select)
                : _graphClient.Users[id].GetAsync(rc => rc.QueryParameters.Select = select);
        }

        public Task<int?> GetUsersRawCountAsync(string filter, string search)
        {
            return _graphClient.Users.Count
               .GetAsync(rc =>
               {
                   rc.Headers = EventualConsistencyHeader;
                   //rc.QueryParameters.Filter = filter;
                   //rc.QueryParameters.Search = WebUtility.UrlEncode(search);
               });

            //LastUrl = WebUtility.UrlDecode(request.GetHttpRequestMessage().RequestUri?.AbsoluteUri);
            //return request.GetAsync();
        }

        public IAsyncEnumerable<Application> GetApplications(string[] select, string filter, string[] orderBy, string search)
        {
            var requestInfo = _graphClient.Applications
                    .CreateGetRequestInformation(rc =>
                    {
                        rc.Headers = EventualConsistencyHeader;
                        rc.QueryParameters.Count = true;
                        rc.QueryParameters.Select = select;
                        rc.QueryParameters.Filter = filter;
                        rc.QueryParameters.Orderby = orderBy;
                        if (!search.IsNullOrEmpty())
                            rc.QueryParameters.Search = search;
                    });

            requestInfo.PathParameters.Add("baseurl", _graphClient.RequestAdapter.BaseUrl); //TODO: remove before GA
            LastUrl = WebUtility.UrlDecode(requestInfo.URI.AbsoluteUri);
            return requestInfo.ToAsyncEnumerable<Application, ApplicationCollectionResponse>(_graphClient.RequestAdapter);
        }

        public Task<ApplicationCollectionResponse> GetApplicationCollectionAsync(string[] select, string filter, string[] orderBy, string search)
        {
            var requestInfo = _graphClient.Applications
                .CreateGetRequestInformation(rc =>
                {
                    rc.Headers = EventualConsistencyHeader;
                    rc.QueryParameters.Count = true;
                    rc.QueryParameters.Select = select;
                    rc.QueryParameters.Filter = filter;
                    rc.QueryParameters.Orderby = orderBy;
                    if (!search.IsNullOrEmpty())
                        rc.QueryParameters.Search = WebUtility.UrlEncode(search);
                });

            requestInfo.PathParameters.Add("baseurl", _graphClient.RequestAdapter.BaseUrl); //TODO: remove before GA
            LastUrl = WebUtility.UrlDecode(requestInfo.URI.AbsoluteUri);
            return _graphClient.RequestAdapter.SendAsync(requestInfo, parseNode => new ApplicationCollectionResponse());

        }

        public IAsyncEnumerable<Device> GetDevices(string[] select, string filter, string[] orderBy, string search)
        {
            var requestInfo = _graphClient.Devices
                    .CreateGetRequestInformation(rc =>
                    {
                        rc.Headers = EventualConsistencyHeader;
                        rc.QueryParameters.Count = true;
                        rc.QueryParameters.Select = select;
                        rc.QueryParameters.Filter = filter;
                        rc.QueryParameters.Orderby = orderBy;
                        if (!search.IsNullOrEmpty())
                            rc.QueryParameters.Search = search;
                    });

            //LastUrl = requestInfo.URI.AbsoluteUri;
            return requestInfo.ToAsyncEnumerable<Device, DeviceCollectionResponse>(_graphClient.RequestAdapter);
        }
        public Task<DeviceCollectionResponse> GetDeviceCollectionAsync(string[] select, string filter, string[] orderBy, string search)
        {
            var requestInfo = _graphClient.Devices
                .CreateGetRequestInformation(rc =>
                {
                    rc.Headers = EventualConsistencyHeader;
                    rc.QueryParameters.Count = true;
                    rc.QueryParameters.Select = select;
                    rc.QueryParameters.Filter = filter;
                    rc.QueryParameters.Orderby = orderBy;
                    if (!search.IsNullOrEmpty())
                        rc.QueryParameters.Search = WebUtility.UrlEncode(search);
                });

            requestInfo.PathParameters.Add("baseurl", _graphClient.RequestAdapter.BaseUrl); //TODO: remove before GA
            LastUrl = WebUtility.UrlDecode(requestInfo.URI.AbsoluteUri);
            return _graphClient.RequestAdapter.SendAsync(requestInfo, parseNode => new DeviceCollectionResponse());
        }


        public IAsyncEnumerable<Group> GetGroups(string[] select, string filter, string[] orderBy, string search)
        {
            var requestInfo = _graphClient.Groups
                    .CreateGetRequestInformation(rc =>
                    {
                        rc.Headers = EventualConsistencyHeader;
                        rc.QueryParameters.Count = true;
                        rc.QueryParameters.Select = select;
                        rc.QueryParameters.Filter = filter;
                        rc.QueryParameters.Orderby = orderBy;
                        if (!search.IsNullOrEmpty())
                            rc.QueryParameters.Search = search;
                    });

            //LastUrl = requestInfo.URI.AbsoluteUri;
            return requestInfo.ToAsyncEnumerable<Group, GroupCollectionResponse>(_graphClient.RequestAdapter);
        }

        public Task<GroupCollectionResponse> GetGroupCollectionAsync(string[] select, string filter, string[] orderBy, string search)
        {
            return _graphClient.Groups
                .GetAsync(rc =>
                {
                    rc.Headers = EventualConsistencyHeader;
                    rc.QueryParameters.Count = true;
                    rc.QueryParameters.Select = select;
                    rc.QueryParameters.Filter = filter;
                    rc.QueryParameters.Orderby = orderBy;
                    if (!search.IsNullOrEmpty())
                        rc.QueryParameters.Search = search;
                });

            //LastUrl = WebUtility.UrlDecode(request.GetHttpRequestMessage().RequestUri?.AbsoluteUri);
            //return request.GetAsync();
        }

        public IAsyncEnumerable<User> GetUsers(string[] select, string filter, string[] orderBy, string search)
        {
            var requestInfo = _graphClient.Users
                .CreateGetRequestInformation(rc =>
                {
                    rc.Headers = EventualConsistencyHeader;
                    rc.QueryParameters.Count = true;
                    rc.QueryParameters.Select = select;
                    rc.QueryParameters.Filter = filter;
                    rc.QueryParameters.Orderby = orderBy;
                    if (!search.IsNullOrEmpty())
                        rc.QueryParameters.Search = search;
                });

            //LastUrl = requestInfo.URI.AbsoluteUri;
            return requestInfo.ToAsyncEnumerable<User, UserCollectionResponse>(_graphClient.RequestAdapter);
        }

        public Task<UserCollectionResponse> GetUserCollectionAsync(string[] select, string filter, string[] orderBy, string search)
        {
            var requestInfo = _graphClient.Users
                .CreateGetRequestInformation(rc =>
                {
                    rc.Headers = EventualConsistencyHeader;
                    rc.QueryParameters.Count = true;
                    rc.QueryParameters.Select = select;
                    rc.QueryParameters.Filter = filter;
                    rc.QueryParameters.Orderby = orderBy;
                    if (!search.IsNullOrEmpty())
                        rc.QueryParameters.Search = WebUtility.UrlEncode(search);
                });

            requestInfo.PathParameters.Add("baseurl", _graphClient.RequestAdapter.BaseUrl); //TODO: remove before GA
            LastUrl = WebUtility.UrlDecode(requestInfo.URI.AbsoluteUri);
            return _graphClient.RequestAdapter.SendAsync(requestInfo, parseNode => new UserCollectionResponse());
        }

        public IAsyncEnumerable<Group> GetTransitiveMemberOfAsGroups(string id)
        {
            var requestInfo = _graphClient.Users[id]
                .TransitiveMemberOf.Group
                .CreateGetRequestInformation(rc =>
                {
                    rc.Headers = EventualConsistencyHeader;
                    rc.QueryParameters.Count = true;
                });

            return requestInfo.ToAsyncEnumerable<Group, GroupCollectionResponse>(_graphClient.RequestAdapter);
        }

        public Task<GroupCollectionResponse> GetTransitiveMemberOfAsGroupCollectionAsync(string id)
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

        public IAsyncEnumerable<User> GetTransitiveMembersAsUsers(string id)
        {
            var requestInfo = _graphClient.Groups[id]
                .TransitiveMembers.User
                .CreateGetRequestInformation(rc =>
                {
                    rc.Headers = EventualConsistencyHeader;
                    rc.QueryParameters.Count = true;
                });

            return requestInfo.ToAsyncEnumerable<User, UserCollectionResponse>(_graphClient.RequestAdapter);
        }

        public Task<UserCollectionResponse> GetTransitiveMembersAsUserCollectionAsync(string id)
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

        public IAsyncEnumerable<User> GetAppOwnersAsUsers(string id)
        {
            var requestInfo = _graphClient.Applications[id]
                .Owners.User
                .CreateGetRequestInformation(rc =>
                {
                    rc.Headers = EventualConsistencyHeader;
                    rc.QueryParameters.Count = true;
                });

            return requestInfo.ToAsyncEnumerable<User, UserCollectionResponse>(_graphClient.RequestAdapter);
        }

        public Task<UserCollectionResponse> GetAppOwnersAsUserCollectionAsync(string id)
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
    }
}