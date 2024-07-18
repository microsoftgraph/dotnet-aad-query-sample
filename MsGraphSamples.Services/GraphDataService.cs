// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Graph.Models.ODataErrors;
using Microsoft.Kiota.Abstractions;
using Microsoft.Kiota.Abstractions.Serialization;
using System.Net;

namespace MsGraphSamples.Services;

public interface IGraphDataService
{
    string? LastUrl { get; }

    Task<TCollectionResponse?> GetNextPageAsync<TCollectionResponse>(TCollectionResponse? collectionResponse) where TCollectionResponse : BaseCollectionPaginationCountResponse, new();

    Task<User?> GetUserAsync(string[] select, string? id = null);
    Task<int?> GetUsersRawCountAsync(string filter, string search);

    Task<ApplicationCollectionResponse?> GetApplicationCollectionAsync(string[] select, string? filter = null, string[]? orderBy = null, string? search = null, ushort top = 999);
    Task<ServicePrincipalCollectionResponse?> GetServicePrincipalCollectionAsync(string[] select, string? filter = null, string[]? orderBy = null, string? search = null, ushort top = 999);
    Task<DeviceCollectionResponse?> GetDeviceCollectionAsync(string[] select, string? filter = null, string[]? orderBy = null, string? search = null, ushort top = 999);
    Task<GroupCollectionResponse?> GetGroupCollectionAsync(string[] select, string? filter = null, string[]? orderBy = null, string? search = null, ushort top = 999);
    Task<UserCollectionResponse?> GetUserCollectionAsync(string[] select, string? filter = null, string[]? orderBy = null, string? search = null, ushort top = 999);

    Task<UserCollectionResponse?> GetApplicationOwnersAsUserCollectionAsync(string id, string[] select, ushort top = 999);
    Task<UserCollectionResponse?> GetServicePrincipalOwnersAsUserCollectionAsync(string id, string[] select, ushort top = 999);
    Task<UserCollectionResponse?> GetDeviceOwnersAsUserCollectionAsync(string id, string[] select, ushort top = 999);
    Task<UserCollectionResponse?> GetGroupOwnersAsUserCollectionAsync(string id, string[] select, ushort top = 999);

    Task<GroupCollectionResponse?> GetTransitiveMemberOfAsGroupCollectionAsync(string id, string[] select, ushort top = 999);
    Task<UserCollectionResponse?> GetTransitiveMembersAsUserCollectionAsync(string id, string[] select, ushort top = 999);

    Task<User?> WriteExtensionProperty(string propertyName, object propertyValue, string userId);
}

public class GraphDataService(GraphServiceClient graphClient) : IGraphDataService
{
    private readonly RequestHeaders EventualConsistencyHeader = new() { { "ConsistencyLevel", "eventual" } };
    private readonly Dictionary<string, ParsableFactory<IParsable>> ErrorMapping = new() { { "XXX", ODataError.CreateFromDiscriminatorValue } };

    public string? LastUrl { get; private set; } = null;

    public Task<User?> GetUserAsync(string[] select, string? id = null)
    {
        return id == null
            ? graphClient.Me.GetAsync(rc => rc.QueryParameters.Select = select)
            : graphClient.Users[id].GetAsync(rc => rc.QueryParameters.Select = select);
    }

    public Task<int?> GetUsersRawCountAsync(string? filter = null, string? search = null)
    {
        var requestInfo = graphClient.Users
            .Count
            .ToGetRequestInformation(rc =>
            {
                rc.Headers = EventualConsistencyHeader;
                rc.QueryParameters.Filter = filter;
                rc.QueryParameters.Search = search;
            });

        LastUrl = WebUtility.UrlDecode(requestInfo.URI.AbsoluteUri);
        return graphClient.RequestAdapter.SendPrimitiveAsync<int?>(requestInfo);
    }

    public Task<TCollectionResponse?> GetNextPageAsync<TCollectionResponse>(TCollectionResponse? collectionResponse) where TCollectionResponse : BaseCollectionPaginationCountResponse, new()
    {
        return collectionResponse.GetNextPageAsync(graphClient.RequestAdapter);
    }

    public Task<ApplicationCollectionResponse?> GetApplicationCollectionAsync(string[] select, string? filter = null, string[]? orderBy = null, string? search = null, ushort top = 999)
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
        return graphClient.RequestAdapter.SendAsync(requestInfo, ApplicationCollectionResponse.CreateFromDiscriminatorValue, ErrorMapping);
    }

    public Task<ServicePrincipalCollectionResponse?> GetServicePrincipalCollectionAsync(string[] select, string? filter = null, string[]? orderBy = null, string? search = null, ushort top = 999)
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
        return graphClient.RequestAdapter.SendAsync(requestInfo, ServicePrincipalCollectionResponse.CreateFromDiscriminatorValue, ErrorMapping);
    }

    public Task<DeviceCollectionResponse?> GetDeviceCollectionAsync(string[] select, string? filter = null, string[]? orderBy = null, string? search = null, ushort top = 999)
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
        return graphClient.RequestAdapter.SendAsync(requestInfo, DeviceCollectionResponse.CreateFromDiscriminatorValue, ErrorMapping);
    }

    public Task<GroupCollectionResponse?> GetGroupCollectionAsync(string[] select, string? filter = null, string[]? orderBy = null, string? search = null, ushort top = 999)
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
        return graphClient.RequestAdapter.SendAsync(requestInfo, GroupCollectionResponse.CreateFromDiscriminatorValue, ErrorMapping);
    }

    public Task<UserCollectionResponse?> GetUserCollectionAsync(string[] select, string? filter = null, string[]? orderBy = null, string? search = null, ushort top = 999)
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
        return graphClient.RequestAdapter.SendAsync(requestInfo, UserCollectionResponse.CreateFromDiscriminatorValue, ErrorMapping);
    }

    public Task<UserCollectionResponse?> GetApplicationOwnersAsUserCollectionAsync(string id, string[] select, ushort top = 999)
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
        return graphClient.RequestAdapter.SendAsync(requestInfo, UserCollectionResponse.CreateFromDiscriminatorValue, ErrorMapping);
    }

    public Task<UserCollectionResponse?> GetServicePrincipalOwnersAsUserCollectionAsync(string id, string[] select, ushort top = 999)
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
        return graphClient.RequestAdapter.SendAsync(requestInfo, UserCollectionResponse.CreateFromDiscriminatorValue, ErrorMapping);
    }

    public Task<UserCollectionResponse?> GetDeviceOwnersAsUserCollectionAsync(string id, string[] select, ushort top = 999)
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
        return graphClient.RequestAdapter.SendAsync(requestInfo, UserCollectionResponse.CreateFromDiscriminatorValue, ErrorMapping);
    }

    public Task<UserCollectionResponse?> GetGroupOwnersAsUserCollectionAsync(string id, string[] select, ushort top = 999)
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
        return graphClient.RequestAdapter.SendAsync(requestInfo, UserCollectionResponse.CreateFromDiscriminatorValue, ErrorMapping);
    }

    public Task<GroupCollectionResponse?> GetTransitiveMemberOfAsGroupCollectionAsync(string id, string[] select, ushort top = 999)
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
        return graphClient.RequestAdapter.SendAsync(requestInfo, GroupCollectionResponse.CreateFromDiscriminatorValue, ErrorMapping);
    }

    public Task<UserCollectionResponse?> GetTransitiveMembersAsUserCollectionAsync(string id, string[] select, ushort top = 999)
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
        return graphClient.RequestAdapter.SendAsync(requestInfo, UserCollectionResponse.CreateFromDiscriminatorValue, ErrorMapping);
    }

    public Task<User?> WriteExtensionProperty(string propertyName, object propertyValue, string userId)
    {
        var userRequestBody = new User();
        userRequestBody.AdditionalData[propertyName] = propertyValue;
        return graphClient.Users[userId].PatchAsync(userRequestBody);
    }
}