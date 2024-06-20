// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using System.Net;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Kiota.Abstractions;

namespace MsGraphSamples.Services;

public interface IAsyncEnumerableGraphDataService
{
    string? LastUrl { get; }
    Task<User?> GetUserAsync(string[] select, string? id = null);
    IAsyncEnumerable<User> GetUsersInBatch(string[] select);
    IAsyncEnumerable<Application> GetApplications(string[] select, string? filter = null, string[]? orderBy = null, string? search = null);
    IAsyncEnumerable<Device> GetDevices(string[] select, string? filter = null, string[]? orderBy = null, string? search = null);
    IAsyncEnumerable<Group> GetGroups(string[] select, string? filter = null, string[]? orderBy = null, string? search = null);
    IAsyncEnumerable<User> GetUsers(string[] select, string? filter = null, string[]? orderBy = null, string? search = null);
    IAsyncEnumerable<Group> GetTransitiveMemberOfAsGroups(string id);
    IAsyncEnumerable<User> GetTransitiveMembersAsUsers(string id);
    IAsyncEnumerable<User> GetAppOwnersAsUsers(string id);
}

public class AsyncEnumerableGraphDataService(GraphServiceClient graphClient) : IAsyncEnumerableGraphDataService
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

    public IAsyncEnumerable<User> GetUsersInBatch(string[] select)
    {
        return _graphClient.Batch<User, UserCollectionResponse>(
            _graphClient.Users.ToGetRequestInformation(rc =>
            {
                rc.QueryParameters.Select = select;
                rc.QueryParameters.Filter = "startsWith(displayName, 'a')";
            }),
            _graphClient.Users.ToGetRequestInformation(rc =>
            {
                rc.QueryParameters.Select = select;
                rc.QueryParameters.Filter = "startsWith(displayName, 'b')";
            }),
            _graphClient.Users.ToGetRequestInformation(rc =>
            {
                rc.QueryParameters.Select = select;
                rc.QueryParameters.Filter = "startsWith(displayName, 'c')";
            }));
    }

    public IAsyncEnumerable<Application> GetApplications(string[] select, string? filter = null, string[]? orderBy = null, string? search = null)
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
                rc.QueryParameters.Top = 999;
            });

        LastUrl = WebUtility.UrlDecode(requestInfo.URI.AbsoluteUri);
        return requestInfo.ToAsyncEnumerable<Application, ApplicationCollectionResponse>(_graphClient.RequestAdapter);
    }

    public IAsyncEnumerable<Device> GetDevices(string[] select, string? filter = null, string[]? orderBy = null, string? search = null)
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
                rc.QueryParameters.Top = 999;
            });

        LastUrl = WebUtility.UrlDecode(requestInfo.URI.AbsoluteUri);
        return requestInfo.ToAsyncEnumerable<Device, DeviceCollectionResponse>(_graphClient.RequestAdapter);
    }

    public IAsyncEnumerable<Group> GetGroups(string[] select, string? filter = null, string[]? orderBy = null, string? search = null)
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
                rc.QueryParameters.Top = 999;
            });

        LastUrl = WebUtility.UrlDecode(requestInfo.URI.AbsoluteUri);
        return requestInfo.ToAsyncEnumerable<Group, GroupCollectionResponse>(_graphClient.RequestAdapter);
    }

    public IAsyncEnumerable<User> GetUsers(string[] select, string? filter = null, string[]? orderBy = null, string? search = null)
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
            rc.QueryParameters.Top = 999;
        });

        LastUrl = WebUtility.UrlDecode(requestInfo.URI.AbsoluteUri);
        return requestInfo.ToAsyncEnumerable<User, UserCollectionResponse>(_graphClient.RequestAdapter);
    }

    public IAsyncEnumerable<Group> GetTransitiveMemberOfAsGroups(string id)
    {
        var requestInfo = _graphClient.Users[id]
            .TransitiveMemberOf.GraphGroup
            .ToGetRequestInformation(rc =>
            {
                rc.Headers = EventualConsistencyHeader;
                rc.QueryParameters.Count = true;
                rc.QueryParameters.Top = 999;
            });

        LastUrl = WebUtility.UrlDecode(requestInfo.URI.AbsoluteUri);
        return requestInfo.ToAsyncEnumerable<Group, GroupCollectionResponse>(_graphClient.RequestAdapter);
    }

    public IAsyncEnumerable<User> GetTransitiveMembersAsUsers(string id)
    {
        var requestInfo = _graphClient.Groups[id]
            .TransitiveMembers.GraphUser
            .ToGetRequestInformation(rc =>
            {
                rc.Headers = EventualConsistencyHeader;
                rc.QueryParameters.Count = true;
                rc.QueryParameters.Top = 999;
            });

        LastUrl = WebUtility.UrlDecode(requestInfo.URI.AbsoluteUri);
        return requestInfo.ToAsyncEnumerable<User, UserCollectionResponse>(_graphClient.RequestAdapter);
    }

    public IAsyncEnumerable<User> GetAppOwnersAsUsers(string id)
    {
        var requestInfo = _graphClient.Applications[id]
            .Owners.GraphUser
            .ToGetRequestInformation(rc =>
            {
                rc.Headers = EventualConsistencyHeader;
                rc.QueryParameters.Count = true;
                rc.QueryParameters.Top = 999;
            });

        LastUrl = WebUtility.UrlDecode(requestInfo.URI.AbsoluteUri);
        return requestInfo.ToAsyncEnumerable<User, UserCollectionResponse>(_graphClient.RequestAdapter);
    }
}