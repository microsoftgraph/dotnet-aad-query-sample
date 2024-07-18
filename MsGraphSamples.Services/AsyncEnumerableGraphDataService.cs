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
    public long? LastCount { get; }

    Task<User?> GetUserAsync(string[] select, string? id = null);
    IAsyncEnumerable<User> GetUsersInBatch(string[] select, ushort top = 999, CancellationToken cancellationToken = default);
    IAsyncEnumerable<Application> GetApplications(string[] select, string? filter = null, string[]? orderBy = null, string? search = null, ushort top = 999, CancellationToken cancellationToken = default);
    IAsyncEnumerable<ServicePrincipal> GetServicePrincipals(string[] splittedSelect, string? filter, string[]? splittedOrderBy, string? search, ushort top = 999, CancellationToken cancellationToken = default);
    IAsyncEnumerable<Device> GetDevices(string[] select, string? filter = null, string[]? orderBy = null, string? search = null, ushort top = 999, CancellationToken cancellationToken = default);
    IAsyncEnumerable<Group> GetGroups(string[] select, string? filter = null, string[]? orderBy = null, string? search = null, ushort top = 999, CancellationToken cancellationToken = default);
    IAsyncEnumerable<User> GetUsers(string[] select, string? filter = null, string[]? orderBy = null, string? search = null, ushort top = 999, CancellationToken cancellationToken = default);
    IAsyncEnumerable<User> GetApplicationOwnersAsUsers(string id, string[] select, ushort top = 999, CancellationToken cancellationToken = default);
    IAsyncEnumerable<User> GetServicePrincipalOwnersAsUsers(string id, string[] select, ushort top = 999, CancellationToken cancellationToken = default);
    IAsyncEnumerable<User> GetDeviceOwnersAsUsers(string id, string[] select, ushort top = 999, CancellationToken cancellationToken = default);
    IAsyncEnumerable<User> GetGroupOwnersAsUsers(string id, string[] select, ushort top = 999, CancellationToken cancellationToken = default);
    IAsyncEnumerable<Group> GetTransitiveMemberOfAsGroups(string id, string[] select, ushort top = 999, CancellationToken cancellationToken = default);
    IAsyncEnumerable<User> GetTransitiveMembersAsUsers(string id, string[] select, ushort top = 999, CancellationToken cancellationToken = default);
}

public class AsyncEnumerableGraphDataService(GraphServiceClient graphClient) : IAsyncEnumerableGraphDataService
{
    private readonly RequestHeaders EventualConsistencyHeader = new() { { "ConsistencyLevel", "eventual" } };

    public string? LastUrl { get; private set; } = null;
    public long? LastCount { get; private set; } = null;
    private void SetCount(long? count) => LastCount = count;

    public Task<User?> GetUserAsync(string[] select, string? id = null)
    {
        return id == null
            ? graphClient.Me.GetAsync(rc => rc.QueryParameters.Select = select)
            : graphClient.Users[id].GetAsync(rc => rc.QueryParameters.Select = select);
    }

    public IAsyncEnumerable<User> GetUsersInBatch(string[] select, ushort top = 999, CancellationToken cancellationToken = default)
    {
        return graphClient.Batch<User, UserCollectionResponse>(cancellationToken,
            graphClient.Users.ToGetRequestInformation(rc =>
            {
                rc.Headers = EventualConsistencyHeader;
                rc.QueryParameters.Count = true;
                rc.QueryParameters.Select = select;
                rc.QueryParameters.Filter = "startsWith(displayName, 'x')";
                rc.QueryParameters.Top = top;
            }),
            graphClient.Users.ToGetRequestInformation(rc =>
            {
                rc.Headers = EventualConsistencyHeader;
                rc.QueryParameters.Count = true;
                rc.QueryParameters.Select = select;
                rc.QueryParameters.Filter = "startsWith(displayName, 'y')";
                rc.QueryParameters.Top = top;
            }),
            graphClient.Users.ToGetRequestInformation(rc =>
            {
                rc.Headers = EventualConsistencyHeader;
                rc.QueryParameters.Count = true;
                rc.QueryParameters.Select = select;
                rc.QueryParameters.Filter = "startsWith(displayName, 'z')";
                rc.QueryParameters.Top = top;
            }));
    }

    public IAsyncEnumerable<Application> GetApplications(string[] select, string? filter = null, string[]? orderBy = null, string? search = null, ushort top = 999, CancellationToken cancellationToken = default)
    {
        var requestInfo = graphClient.Applications
            .ToGetRequestInformation(rc =>
            {
                rc.Headers = EventualConsistencyHeader;
                rc.QueryParameters.Count = true;
                rc.QueryParameters.Select = select;
                rc.QueryParameters.Filter = filter;
                rc.QueryParameters.Orderby = orderBy;
                rc.QueryParameters.Search = search;
                rc.QueryParameters.Top = top;
            });

        LastUrl = WebUtility.UrlDecode(requestInfo.URI.AbsoluteUri);
        return requestInfo.ToAsyncEnumerable<Application, ApplicationCollectionResponse>(graphClient.RequestAdapter, SetCount, cancellationToken);
    }

    public IAsyncEnumerable<ServicePrincipal> GetServicePrincipals(string[] select, string? filter, string[]? orderBy, string? search = null, ushort top = 999, CancellationToken cancellationToken = default)
    {
        var requestInfo = graphClient.ServicePrincipals
            .ToGetRequestInformation(rc =>
            {
                rc.Headers = EventualConsistencyHeader;
                rc.QueryParameters.Count = true;
                rc.QueryParameters.Select = select;
                rc.QueryParameters.Filter = filter;
                rc.QueryParameters.Orderby = orderBy;
                rc.QueryParameters.Search = search;
                rc.QueryParameters.Top = top;
            });

        LastUrl = WebUtility.UrlDecode(requestInfo.URI.AbsoluteUri);
        return requestInfo.ToAsyncEnumerable<ServicePrincipal, ServicePrincipalCollectionResponse>(graphClient.RequestAdapter, SetCount, cancellationToken);
    }

    public IAsyncEnumerable<Device> GetDevices(string[] select, string? filter = null, string[]? orderBy = null, string? search = null, ushort top = 999, CancellationToken cancellationToken = default)
    {
        var requestInfo = graphClient.Devices
            .ToGetRequestInformation(rc =>
            {
                rc.Headers = EventualConsistencyHeader;
                rc.QueryParameters.Count = true;
                rc.QueryParameters.Select = select;
                rc.QueryParameters.Filter = filter;
                rc.QueryParameters.Orderby = orderBy;
                rc.QueryParameters.Search = search;
                rc.QueryParameters.Top = top;
            });

        LastUrl = WebUtility.UrlDecode(requestInfo.URI.AbsoluteUri);
        return requestInfo.ToAsyncEnumerable<Device, DeviceCollectionResponse>(graphClient.RequestAdapter, SetCount, cancellationToken);
    }

    public IAsyncEnumerable<Group> GetGroups(string[] select, string? filter = null, string[]? orderBy = null, string? search = null, ushort top = 999, CancellationToken cancellationToken = default)
    {
        var requestInfo = graphClient.Groups
            .ToGetRequestInformation(rc =>
            {
                rc.Headers = EventualConsistencyHeader;
                rc.QueryParameters.Count = true;
                rc.QueryParameters.Select = select;
                rc.QueryParameters.Filter = filter;
                rc.QueryParameters.Orderby = orderBy;
                rc.QueryParameters.Search = search;
                rc.QueryParameters.Top = top;
            });

        LastUrl = WebUtility.UrlDecode(requestInfo.URI.AbsoluteUri);
        return requestInfo.ToAsyncEnumerable<Group, GroupCollectionResponse>(graphClient.RequestAdapter, SetCount, cancellationToken);
    }

    public IAsyncEnumerable<User> GetUsers(string[] select, string? filter = null, string[]? orderBy = null, string? search = null, ushort top = 999, CancellationToken cancellationToken = default)
    {
        var requestInfo = graphClient.Users
        .ToGetRequestInformation(rc =>
        {
            rc.Headers = EventualConsistencyHeader;
            rc.QueryParameters.Count = true;
            rc.QueryParameters.Select = select;
            rc.QueryParameters.Filter = filter;
            rc.QueryParameters.Orderby = orderBy;
            rc.QueryParameters.Search = search;
            rc.QueryParameters.Top = top;
        });

        LastUrl = WebUtility.UrlDecode(requestInfo.URI.AbsoluteUri);
        return requestInfo.ToAsyncEnumerable<User, UserCollectionResponse>(graphClient.RequestAdapter, SetCount, cancellationToken);
    }

    public IAsyncEnumerable<User> GetApplicationOwnersAsUsers(string id, string[] select, ushort top = 999, CancellationToken cancellationToken = default)
    {
        var requestInfo = graphClient.Applications[id]
            .Owners.GraphUser
            .ToGetRequestInformation(rc =>
            {
                rc.Headers = EventualConsistencyHeader;
                rc.QueryParameters.Count = true;
                rc.QueryParameters.Select = select;
                rc.QueryParameters.Top = top;
            });

        LastUrl = WebUtility.UrlDecode(requestInfo.URI.AbsoluteUri);
        return requestInfo.ToAsyncEnumerable<User, UserCollectionResponse>(graphClient.RequestAdapter, SetCount, cancellationToken);
    }

    public IAsyncEnumerable<User> GetServicePrincipalOwnersAsUsers(string id, string[] select, ushort top = 999, CancellationToken cancellationToken = default)
    {
        var requestInfo = graphClient.ServicePrincipals[id]
            .Owners.GraphUser
            .ToGetRequestInformation(rc =>
            {
                rc.Headers = EventualConsistencyHeader;
                rc.QueryParameters.Count = true;
                rc.QueryParameters.Select = select;
                rc.QueryParameters.Top = top;
            });

        LastUrl = WebUtility.UrlDecode(requestInfo.URI.AbsoluteUri);
        return requestInfo.ToAsyncEnumerable<User, UserCollectionResponse>(graphClient.RequestAdapter, SetCount, cancellationToken);
    }

    public IAsyncEnumerable<User> GetDeviceOwnersAsUsers(string id, string[] select, ushort top = 999, CancellationToken cancellationToken = default)
    {
        var requestInfo = graphClient.Devices[id]
            .RegisteredOwners.GraphUser
            .ToGetRequestInformation(rc =>
            {
                rc.Headers = EventualConsistencyHeader;
                rc.QueryParameters.Count = true;
                rc.QueryParameters.Select = select;
                rc.QueryParameters.Top = top;
            });

        LastUrl = WebUtility.UrlDecode(requestInfo.URI.AbsoluteUri);
        return requestInfo.ToAsyncEnumerable<User, UserCollectionResponse>(graphClient.RequestAdapter, SetCount, cancellationToken);
    }

    public IAsyncEnumerable<User> GetGroupOwnersAsUsers(string id, string[] select, ushort top = 999, CancellationToken cancellationToken = default)
    {
        var requestInfo = graphClient.Groups[id]
            .Owners.GraphUser
            .ToGetRequestInformation(rc =>
            {

                rc.Headers = EventualConsistencyHeader;
                rc.QueryParameters.Count = true;
                rc.QueryParameters.Select = select;
                rc.QueryParameters.Top = top;
            });

        LastUrl = WebUtility.UrlDecode(requestInfo.URI.AbsoluteUri);
        return requestInfo.ToAsyncEnumerable<User, UserCollectionResponse>(graphClient.RequestAdapter, SetCount, cancellationToken);
    }

    public IAsyncEnumerable<Group> GetTransitiveMemberOfAsGroups(string id, string[] select, ushort top = 999, CancellationToken cancellationToken = default)
    {
        var requestInfo = graphClient.Users[id]
            .TransitiveMemberOf.GraphGroup
            .ToGetRequestInformation(rc =>
            {
                rc.Headers = EventualConsistencyHeader;
                rc.QueryParameters.Count = true;
                rc.QueryParameters.Select = select;
                rc.QueryParameters.Top = top;
            });

        LastUrl = WebUtility.UrlDecode(requestInfo.URI.AbsoluteUri);
        return requestInfo.ToAsyncEnumerable<Group, GroupCollectionResponse>(graphClient.RequestAdapter, SetCount, cancellationToken);
    }

    public IAsyncEnumerable<User> GetTransitiveMembersAsUsers(string id, string[] select, ushort top = 999, CancellationToken cancellationToken = default)
    {
        var requestInfo = graphClient.Groups[id]
            .TransitiveMembers.GraphUser
            .ToGetRequestInformation(rc =>
            {
                rc.Headers = EventualConsistencyHeader;
                rc.QueryParameters.Count = true;
                rc.QueryParameters.Select = select;
                rc.QueryParameters.Top = top;
            });

        LastUrl = WebUtility.UrlDecode(requestInfo.URI.AbsoluteUri);
        return requestInfo.ToAsyncEnumerable<User, UserCollectionResponse>(graphClient.RequestAdapter, SetCount, cancellationToken);
    }
}