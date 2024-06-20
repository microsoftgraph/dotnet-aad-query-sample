// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Kiota.Abstractions;
using System.Net;

namespace MsGraphSamples.Services;
public interface IGraphDataService
{
    string? LastUrl { get; }

    Task<User?> GetUserAsync(string[] select, string? id = null);
    Task<int?> GetUsersRawCountAsync(string filter, string search);

    Task<ApplicationCollectionResponse?> GetApplicationCollectionAsync(string[] select, string? filter = null, string[]? orderBy = null, string? search = null);

    Task<DeviceCollectionResponse?> GetDeviceCollectionAsync(string[] select, string? filter = null, string[]? orderBy = null, string? search = null);

    Task<GroupCollectionResponse?> GetGroupCollectionAsync(string[] select, string? filter = null, string[]? orderBy = null, string? search = null);

    Task<UserCollectionResponse?> GetUserCollectionAsync(string[] select, string? filter = null, string[]? orderBy = null, string? search = null);

    Task<GroupCollectionResponse?> GetTransitiveMemberOfAsGroupCollectionAsync(string id);

    Task<UserCollectionResponse?> GetTransitiveMembersAsUserCollectionAsync(string id);

    Task<UserCollectionResponse?> GetAppOwnersAsUserCollectionAsync(string id);
    Task WriteExtensionProperty(string propertyName, object propertyValue, string userId);
}

public class GraphDataService(GraphServiceClient graphClient) : IGraphDataService
{
    private readonly GraphServiceClient _graphClient = graphClient;

    private readonly RequestHeaders EventualConsistencyHeader = new() { { "ConsistencyLevel", "eventual" } };

    public string? LastUrl { get; private set; } = null;

    public Task<User?> GetUserAsync(string[] select, string? id = null)
    {
        return id == null
            ? _graphClient.Me.GetAsync(rc => rc.QueryParameters.Select = select)
            : _graphClient.Users[id].GetAsync(rc => rc.QueryParameters.Select = select);
    }

    public async Task WriteExtensionProperty(string propertyName, object propertyValue, string userId)
    {
        var userRequestBody = new User();
        userRequestBody.AdditionalData[propertyName] = propertyValue;
        await _graphClient.Users[userId].PatchAsync(userRequestBody);
    }

    public Task<int?> GetUsersRawCountAsync(string? filter = null, string? search = null)
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

    public Task<ApplicationCollectionResponse?> GetApplicationCollectionAsync(string[] select, string? filter = null, string[]? orderBy = null, string? search = null)
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

    public Task<DeviceCollectionResponse?> GetDeviceCollectionAsync(string[] select, string? filter = null, string[]? orderBy = null, string? search = null)
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

    public Task<GroupCollectionResponse?> GetGroupCollectionAsync(string[] select, string? filter = null, string[]? orderBy = null, string? search = null)
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

    public Task<UserCollectionResponse?> GetUserCollectionAsync(string[] select, string? filter = null, string[]? orderBy = null, string? search = null)
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

    public Task<GroupCollectionResponse?> GetTransitiveMemberOfAsGroupCollectionAsync(string id)
    {
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

    public Task<UserCollectionResponse?> GetTransitiveMembersAsUserCollectionAsync(string id)
    {
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

    public Task<UserCollectionResponse?> GetAppOwnersAsUserCollectionAsync(string id)
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