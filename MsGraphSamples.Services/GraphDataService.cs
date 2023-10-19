// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Kiota.Abstractions;
using System.Net;

namespace MsGraphSamples.Services;

public interface IGraphDataService
{
    string? LastUrl
    {
        get;
    }
    Task<TCollectionResponse?> GetNextPageAsync<TCollectionResponse>(TCollectionResponse collectionResponse)
        where TCollectionResponse : BaseCollectionPaginationCountResponse, new();
    
    Task<User?> GetUserAsync(string[] select, string? id = null);
    Task<int?> GetUsersRawCountAsync(string filter, string search);

    Task<ApplicationCollectionResponse?> GetApplicationCollectionAsync(string[] select, string? filter = null, string[]? orderBy = null, string? search = null);

    Task<ServicePrincipalCollectionResponse?> GetServicePrincipalsCollectionAsync(string[] select, string? filter = null, string[]? orderBy = null, string? search = null);

    Task<DeviceCollectionResponse?> GetDeviceCollectionAsync(string[] select, string? filter = null, string[]? orderBy = null, string? search = null);

    Task<GroupCollectionResponse?> GetGroupCollectionAsync(string[] select, string? filter = null, string[]? orderBy = null, string? search = null);

    Task<UserCollectionResponse?> GetUserCollectionAsync(string[] select, string? filter = null, string[]? orderBy = null, string? search = null);

    Task<GroupCollectionResponse?> GetTransitiveMemberOfAsGroupCollectionAsync(string id, string[] select);

    Task<UserCollectionResponse?> GetTransitiveMembersAsUserCollectionAsync(string id, string[] select);

    Task<UserCollectionResponse?> GetAppOwnersAsUserCollectionAsync(string id, string[] select);
    Task<UserCollectionResponse?> GetSPOwnersAsUserCollectionAsync(string id, string[] select);
    Task<UserCollectionResponse?> GetRegisteredOwnersAsUserCollectionAsync(string id, string[] select);

    Task WriteExtensionProperty(string propertyName, object propertyValue, string userId);
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

    public Task<TCollectionResponse?> GetNextPageAsync<TCollectionResponse>(TCollectionResponse collectionResponse)
           where TCollectionResponse : BaseCollectionPaginationCountResponse, new()
    {
        if (collectionResponse.OdataNextLink == null)
        {
            return Task.FromResult<TCollectionResponse?>(null);
        }

        var nextPageRequestInformation = new RequestInformation
        {
            HttpMethod = Method.GET,
            UrlTemplate = collectionResponse.OdataNextLink,
        };

        return _graphClient.RequestAdapter
            .SendAsync(nextPageRequestInformation, parseNode => new TCollectionResponse());
    }

    public Task<User?> GetUserAsync(string[] select, string? id = null)
    {
        return id == null
            ? _graphClient.Me.GetAsync(rc => rc.QueryParameters.Select = select)
            : _graphClient.Users[id].GetAsync(rc => rc.QueryParameters.Select = select);
    }


    public  Task WriteExtensionProperty(string propertyName, object propertyValue, string userId)
    {
        var userRequestBody = new User();
        userRequestBody.AdditionalData[propertyName] = propertyValue;
        return _graphClient.Users[userId].PatchAsync(userRequestBody);
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
    public Task<ServicePrincipalCollectionResponse?> GetServicePrincipalsCollectionAsync(string[] select, string? filter = null, string[]? orderBy = null, string? search = null)
    {
        var requestInfo = _graphClient.ServicePrincipals
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
        return _graphClient.RequestAdapter.SendAsync(requestInfo, ServicePrincipalCollectionResponse.CreateFromDiscriminatorValue);
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

    public Task<GroupCollectionResponse?> GetTransitiveMemberOfAsGroupCollectionAsync(string id, string[] select)
    {
        ArgumentNullException.ThrowIfNull(id);

        var requestInfo = _graphClient.Users[id]
            .TransitiveMemberOf.GraphGroup
            .ToGetRequestInformation(rc =>
            {
                rc.Headers = EventualConsistencyHeader;
                rc.QueryParameters.Count = true;
                rc.QueryParameters.Select = select;
            });

        LastUrl = WebUtility.UrlDecode(requestInfo.URI.AbsoluteUri);
        return _graphClient.RequestAdapter.SendAsync(requestInfo, GroupCollectionResponse.CreateFromDiscriminatorValue);
    }

    public Task<UserCollectionResponse?> GetTransitiveMembersAsUserCollectionAsync(string id, string[] select)
    {
        ArgumentNullException.ThrowIfNull(id);

        var requestInfo = _graphClient.Groups[id]
            .TransitiveMembers.GraphUser
            .ToGetRequestInformation(rc =>
            {
                rc.Headers = EventualConsistencyHeader;
                rc.QueryParameters.Count = true;
                rc.QueryParameters.Select = select;
            });

        LastUrl = WebUtility.UrlDecode(requestInfo.URI.AbsoluteUri);
        return _graphClient.RequestAdapter.SendAsync(requestInfo, UserCollectionResponse.CreateFromDiscriminatorValue);
    }

    public Task<UserCollectionResponse?> GetAppOwnersAsUserCollectionAsync(string id, string[] select)
    {
        var requestInfo = _graphClient.Applications[id]
            .Owners.GraphUser
            .ToGetRequestInformation(rc =>
            {
                rc.Headers = EventualConsistencyHeader;
                rc.QueryParameters.Count = true;
                rc.QueryParameters.Select = select;
            });

        LastUrl = WebUtility.UrlDecode(requestInfo.URI.AbsoluteUri);
        return _graphClient.RequestAdapter.SendAsync(requestInfo, UserCollectionResponse.CreateFromDiscriminatorValue);
    }

    public Task<UserCollectionResponse?> GetSPOwnersAsUserCollectionAsync(string id, string[] select)
    {
        var requestInfo = _graphClient.ServicePrincipals[id]
            .Owners.GraphUser
            .ToGetRequestInformation(rc =>
            {
                rc.Headers = EventualConsistencyHeader;
                rc.QueryParameters.Count = true;
                rc.QueryParameters.Select = select;
            });

        LastUrl = WebUtility.UrlDecode(requestInfo.URI.AbsoluteUri);
        return _graphClient.RequestAdapter.SendAsync(requestInfo, UserCollectionResponse.CreateFromDiscriminatorValue);
    }

    public Task<UserCollectionResponse?> GetRegisteredOwnersAsUserCollectionAsync(string id, string[] select)
    {
        var requestInfo = _graphClient.Devices[id]
            .RegisteredOwners.GraphUser
            .ToGetRequestInformation(rc =>
            {
                rc.Headers = EventualConsistencyHeader;
                rc.QueryParameters.Count = true;
                rc.QueryParameters.Select = select;
            });

        LastUrl = WebUtility.UrlDecode(requestInfo.URI.AbsoluteUri);
        return _graphClient.RequestAdapter.SendAsync(requestInfo, UserCollectionResponse.CreateFromDiscriminatorValue);
    }
}