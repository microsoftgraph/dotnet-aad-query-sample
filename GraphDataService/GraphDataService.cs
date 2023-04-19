// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Kiota.Abstractions;
using MsGraph_Samples.Helpers;
using System.Net;

namespace MsGraph_Samples.Services
{
    public interface IGraphDataService
    {
        string? LastUrl { get; }
        Task<User?> GetUserAsync(string[] select, string? id = null);
        Task<int?> GetUsersRawCountAsync(string filter, string search);

        Task<(int membersCount, int guestsCount)> BatchRequest();

        IAsyncEnumerable<Application> GetApplications(string[] select, string filter, string[] orderBy, string search);
        Task<ApplicationCollectionResponse?> GetApplicationCollectionAsync(string[] select, string filter, string[] orderBy, string search);

        IAsyncEnumerable<Device> GetDevices(string[] select, string filter, string[] orderBy, string search);
        Task<DeviceCollectionResponse?> GetDeviceCollectionAsync(string[] select, string filter, string[] orderBy, string search);

        IAsyncEnumerable<Group> GetGroups(string[] select, string filter, string[] orderBy, string search);
        Task<GroupCollectionResponse?> GetGroupCollectionAsync(string[] select, string filter, string[] orderBy, string search);

        IAsyncEnumerable<User> GetUsers(string[] select, string filter, string[] orderBy, string search);
        Task<UserCollectionResponse?> GetUserCollectionAsync(string[] select, string filter, string[] orderBy, string search);

        IAsyncEnumerable<Group> GetTransitiveMemberOfAsGroups(string? id);
        Task<GroupCollectionResponse?> GetTransitiveMemberOfAsGroupCollectionAsync(string? id);

        IAsyncEnumerable<User> GetTransitiveMembersAsUsers(string? id);
        Task<UserCollectionResponse?> GetTransitiveMembersAsUserCollectionAsync(string? id);

        IAsyncEnumerable<User> GetAppOwnersAsUsers(string? id);
        Task<UserCollectionResponse?> GetAppOwnersAsUserCollectionAsync(string? id);
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

        public Task<User?> GetUserAsync(string[] select, string? id = null)
        {
            return id == null
                ? _graphClient.Me.GetAsync(rc => rc.QueryParameters.Select = select)
                : _graphClient.Users[id].GetAsync(rc => rc.QueryParameters.Select = select);
        }

        public async Task<(int membersCount, int guestsCount)> BatchRequest()
        {
            var membersCountRequest = _graphClient.Users.Count
                .ToGetRequestInformation(rc =>
                {
                    rc.QueryParameters.Filter = "userType eq 'Member'";
                    rc.Headers = EventualConsistencyHeader;
                });
            var guestsCountRequest = _graphClient.Users.Count
                .ToGetRequestInformation(rc =>
                {
                    rc.QueryParameters.Filter = "userType eq 'Guest'";
                    rc.Headers = EventualConsistencyHeader;
                });

            var batchRequestContent = new BatchRequestContent(_graphClient);

            var membersCountRequestId = await batchRequestContent.AddBatchRequestStepAsync(membersCountRequest);
            var guestsCountRequestId = await batchRequestContent.AddBatchRequestStepAsync(guestsCountRequest);

            var batchResponse = await _graphClient.Batch.PostAsync(batchRequestContent);

            var membersCountResponse = await batchResponse.GetResponseByIdAsync(membersCountRequestId);
            var guestsCountResponse = await batchResponse.GetResponseByIdAsync(guestsCountRequestId);

            return (
                Convert.ToInt32(await membersCountResponse.Content.ReadAsStringAsync()),
                Convert.ToInt32(await guestsCountResponse.Content.ReadAsStringAsync()));
        }


        public Task<int?> GetUsersRawCountAsync(string filter, string search)
        {
            var requestInfo = _graphClient.Users.Count
               .ToGetRequestInformation(rc =>
               {
                   rc.Headers = EventualConsistencyHeader;
                   rc.QueryParameters.Filter = filter;
                   rc.QueryParameters.Search = search;
               });

            LastUrl = WebUtility.UrlDecode(requestInfo.URI.AbsoluteUri);
            return _graphClient.RequestAdapter.SendPrimitiveAsync<int?>(requestInfo);
        }

        public IAsyncEnumerable<Application> GetApplications(string[] select, string filter, string[] orderBy, string search)
        {
            var requestInfo = _graphClient.Applications
                    .ToGetRequestInformation(rc =>
                    {
                        rc.Headers = EventualConsistencyHeader;
                        rc.QueryParameters.Count = true;
                        rc.QueryParameters.Select = select;
                        rc.QueryParameters.Filter = filter;
                        rc.QueryParameters.Orderby = orderBy;
                        rc.QueryParameters.Search = search;
                    });

            LastUrl = WebUtility.UrlDecode(requestInfo.URI.AbsoluteUri);
            return requestInfo.ToAsyncEnumerable<Application, ApplicationCollectionResponse>(_graphClient.RequestAdapter);
        }

        public Task<ApplicationCollectionResponse?> GetApplicationCollectionAsync(string[] select, string filter, string[] orderBy, string search)
        {
            var requestInfo = _graphClient.Applications
                .ToGetRequestInformation(rc =>
                {
                    rc.Headers = EventualConsistencyHeader;
                    rc.QueryParameters.Count = true;
                    rc.QueryParameters.Select = select;
                    rc.QueryParameters.Filter = filter;
                    rc.QueryParameters.Orderby = orderBy;
                    rc.QueryParameters.Search = search;
                });

            LastUrl = WebUtility.UrlDecode(requestInfo.URI.AbsoluteUri);
            return _graphClient.RequestAdapter.SendAsync(requestInfo, ApplicationCollectionResponse.CreateFromDiscriminatorValue);
        }

        public IAsyncEnumerable<Device> GetDevices(string[] select, string filter, string[] orderBy, string search)
        {
            var requestInfo = _graphClient.Devices
                    .ToGetRequestInformation(rc =>
                    {
                        rc.Headers = EventualConsistencyHeader;
                        rc.QueryParameters.Count = true;
                        rc.QueryParameters.Select = select;
                        rc.QueryParameters.Filter = filter;
                        rc.QueryParameters.Orderby = orderBy;
                        rc.QueryParameters.Search = search;
                    });

            LastUrl = WebUtility.UrlDecode(requestInfo.URI.AbsoluteUri);
            return requestInfo.ToAsyncEnumerable<Device, DeviceCollectionResponse>(_graphClient.RequestAdapter);
        }
        public Task<DeviceCollectionResponse?> GetDeviceCollectionAsync(string[] select, string filter, string[] orderBy, string search)
        {
            var requestInfo = _graphClient.Devices
                .ToGetRequestInformation(rc =>
                {
                    rc.Headers = EventualConsistencyHeader;
                    rc.QueryParameters.Count = true;
                    rc.QueryParameters.Select = select;
                    rc.QueryParameters.Filter = filter;
                    rc.QueryParameters.Orderby = orderBy;
                    rc.QueryParameters.Search = search;
                });

            LastUrl = WebUtility.UrlDecode(requestInfo.URI.AbsoluteUri);
            return _graphClient.RequestAdapter.SendAsync(requestInfo, DeviceCollectionResponse.CreateFromDiscriminatorValue);
        }


        public IAsyncEnumerable<Group> GetGroups(string[] select, string filter, string[] orderBy, string search)
        {
            var requestInfo = _graphClient.Groups
                    .ToGetRequestInformation(rc =>
                    {
                        rc.Headers = EventualConsistencyHeader;
                        rc.QueryParameters.Count = true;
                        rc.QueryParameters.Select = select;
                        rc.QueryParameters.Filter = filter;
                        rc.QueryParameters.Orderby = orderBy;
                        rc.QueryParameters.Search = search;
                    });

            LastUrl = WebUtility.UrlDecode(requestInfo.URI.AbsoluteUri);
            return requestInfo.ToAsyncEnumerable<Group, GroupCollectionResponse>(_graphClient.RequestAdapter);
        }

        public Task<GroupCollectionResponse?> GetGroupCollectionAsync(string[] select, string filter, string[] orderBy, string search)
        {
            var requestInfo = _graphClient.Groups
                .ToGetRequestInformation(rc =>
                {
                    rc.Headers = EventualConsistencyHeader;
                    rc.QueryParameters.Count = true;
                    rc.QueryParameters.Select = select;
                    rc.QueryParameters.Filter = filter;
                    rc.QueryParameters.Orderby = orderBy;
                    rc.QueryParameters.Search = search;
                });

            LastUrl = WebUtility.UrlDecode(requestInfo.URI.AbsoluteUri);
            return _graphClient.RequestAdapter.SendAsync(requestInfo, GroupCollectionResponse.CreateFromDiscriminatorValue);
        }

        public IAsyncEnumerable<User> GetUsers(string[] select, string filter, string[] orderBy, string search)
        {
            var requestInfo = _graphClient.Users
                .ToGetRequestInformation(rc =>
                {
                    rc.Headers = EventualConsistencyHeader;
                    rc.QueryParameters.Count = true;
                    rc.QueryParameters.Select = select;
                    rc.QueryParameters.Filter = filter;
                    rc.QueryParameters.Orderby = orderBy;
                    rc.QueryParameters.Search = search;
                });

            LastUrl = WebUtility.UrlDecode(requestInfo.URI.AbsoluteUri);
            return requestInfo.ToAsyncEnumerable<User, UserCollectionResponse>(_graphClient.RequestAdapter);
        }

        public Task<UserCollectionResponse?> GetUserCollectionAsync(string[] select, string filter, string[] orderBy, string search)
        {
            var requestInfo = _graphClient.Users
                .ToGetRequestInformation(rc =>
                {
                    rc.Headers = EventualConsistencyHeader;
                    rc.QueryParameters.Count = true;
                    rc.QueryParameters.Select = select;
                    rc.QueryParameters.Filter = filter;
                    rc.QueryParameters.Orderby = orderBy;
                    rc.QueryParameters.Search = search;
                });

            LastUrl = WebUtility.UrlDecode(requestInfo.URI.AbsoluteUri);
            return _graphClient.RequestAdapter.SendAsync(requestInfo, UserCollectionResponse.CreateFromDiscriminatorValue);
        }

        public IAsyncEnumerable<Group> GetTransitiveMemberOfAsGroups(string? id)
        {
            ArgumentNullException.ThrowIfNull(id);

            var requestInfo = _graphClient.Users[id]
                .TransitiveMemberOf.GraphGroup
                .ToGetRequestInformation(rc =>
                {
                    rc.Headers = EventualConsistencyHeader;
                    rc.QueryParameters.Count = true;
                });

            LastUrl = WebUtility.UrlDecode(requestInfo.URI.AbsoluteUri);
            return requestInfo.ToAsyncEnumerable<Group, GroupCollectionResponse>(_graphClient.RequestAdapter);
        }

        public Task<GroupCollectionResponse?> GetTransitiveMemberOfAsGroupCollectionAsync(string? id)
        {
            ArgumentNullException.ThrowIfNull(id);

            var requestInfo = _graphClient.Users[id]
                .TransitiveMemberOf.GraphGroup
                .ToGetRequestInformation(rc =>
                {
                    rc.Headers = EventualConsistencyHeader;
                    rc.QueryParameters.Count = true;
                });

            LastUrl = WebUtility.UrlDecode(requestInfo.URI.AbsoluteUri);
            return _graphClient.RequestAdapter.SendAsync(requestInfo, GroupCollectionResponse.CreateFromDiscriminatorValue);
        }

        public IAsyncEnumerable<User> GetTransitiveMembersAsUsers(string? id)
        {
            ArgumentNullException.ThrowIfNull(id);

            var requestInfo = _graphClient.Groups[id]
                .TransitiveMembers.GraphUser
                .ToGetRequestInformation(rc =>
                {
                    rc.Headers = EventualConsistencyHeader;
                    rc.QueryParameters.Count = true;
                });

            LastUrl = WebUtility.UrlDecode(requestInfo.URI.AbsoluteUri);
            return requestInfo.ToAsyncEnumerable<User, UserCollectionResponse>(_graphClient.RequestAdapter);
        }

        public Task<UserCollectionResponse?> GetTransitiveMembersAsUserCollectionAsync(string? id)
        {
            ArgumentNullException.ThrowIfNull(id);

            var requestInfo = _graphClient.Groups[id]
                .TransitiveMembers.GraphUser
                .ToGetRequestInformation(rc =>
                {
                    rc.Headers = EventualConsistencyHeader;
                    rc.QueryParameters.Count = true;
                });

            LastUrl = WebUtility.UrlDecode(requestInfo.URI.AbsoluteUri);
            return _graphClient.RequestAdapter.SendAsync(requestInfo, UserCollectionResponse.CreateFromDiscriminatorValue);
        }

        public IAsyncEnumerable<User> GetAppOwnersAsUsers(string? id)
        {
            ArgumentNullException.ThrowIfNull(id);

            var requestInfo = _graphClient.Applications[id]
                .Owners.GraphUser
                .ToGetRequestInformation(rc =>
                {
                    rc.Headers = EventualConsistencyHeader;
                    rc.QueryParameters.Count = true;
                });

            LastUrl = WebUtility.UrlDecode(requestInfo.URI.AbsoluteUri);
            return requestInfo.ToAsyncEnumerable<User, UserCollectionResponse>(_graphClient.RequestAdapter);
        }

        public Task<UserCollectionResponse?> GetAppOwnersAsUserCollectionAsync(string? id)
        {
            var requestInfo = _graphClient.Applications[id]
                .Owners.GraphUser
                .ToGetRequestInformation(rc =>
                {
                    rc.Headers = EventualConsistencyHeader;
                    rc.QueryParameters.Count = true;
                });

            LastUrl = WebUtility.UrlDecode(requestInfo.URI.AbsoluteUri);
            return _graphClient.RequestAdapter.SendAsync(requestInfo, UserCollectionResponse.CreateFromDiscriminatorValue);
        }
    }
}