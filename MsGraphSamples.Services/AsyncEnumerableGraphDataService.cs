// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using System.Net;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Kiota.Abstractions;
using Microsoft.Kiota.Abstractions.Serialization;

namespace MsGraphSamples.Services;

public static class IAsyncEnumerableGraphExtensions
{
    /// <summary>
    /// Transform a generic RequestInformation into an AsyncEnumerable to efficiently iterate through the collection in case there are several pages.
    /// </summary>
    /// <param name="requestInfo"></param>
    /// <param name="requestAdapter"></param>
    /// <returns>IAsyncEnumerable<User></returns>
    public static async IAsyncEnumerable<TEntity> ToAsyncEnumerable<TEntity, TCollectionResponse>(this RequestInformation requestInfo, IRequestAdapter requestAdapter)
        where TEntity : Entity
        where TCollectionResponse : BaseCollectionPaginationCountResponse, new()
    {
        var collectionResponse = await requestAdapter
        .SendAsync(requestInfo, parseNode => new TCollectionResponse())
        .ConfigureAwait(false);

        await foreach (var entity in collectionResponse.ToAsyncEnumerable<TEntity, TCollectionResponse>(requestAdapter))
        {
            yield return entity;
        }
    }

    public static async IAsyncEnumerable<TEntity> ToAsyncEnumerable<TEntity, TCollectionResponse>(this TCollectionResponse? collectionResponse, IRequestAdapter requestAdapter)
        where TEntity : Entity
        where TCollectionResponse : BaseCollectionPaginationCountResponse, new()
    {
        while (true)
        {
            if (collectionResponse?.GetType().GetProperty("Value")?.GetValue(collectionResponse) is not List<TEntity> entities)
            {
                // not a collection response
                break;
            }

            foreach (var entity in entities)
            {
                yield return entity;
            }

            if (collectionResponse?.OdataNextLink == null)
            {
                break;
            }

            var nextPageRequestInformation = new RequestInformation
            {
                HttpMethod = Method.GET,
                UrlTemplate = collectionResponse.OdataNextLink,
            };

            collectionResponse = await requestAdapter
                .SendAsync(nextPageRequestInformation, parseNode => new TCollectionResponse())
                .ConfigureAwait(false);
        }
    }

    public static async IAsyncEnumerable<TEntity> Batch<TEntity, TCollectionResponse>(this GraphServiceClient graphClient, params RequestInformation[] requests)
        where TEntity : Entity
        where TCollectionResponse : BaseCollectionPaginationCountResponse, new()
    {
        await foreach (var response in graphClient.Batch<TCollectionResponse>(requests))
        {
            await foreach (var entity in response.ToAsyncEnumerable<TEntity, TCollectionResponse>(graphClient.RequestAdapter))
            {
                yield return entity;
            }
        }
    }

    public static async IAsyncEnumerable<T> Batch<T>(this GraphServiceClient graphClient, params RequestInformation[] requests)
        where T : IParsable, new()
    {
        BatchRequestContent batchRequestContent = new(graphClient);

        var addBatchTasks = requests.Select(batchRequestContent.AddBatchRequestStepAsync);
        var requestIds = await Task.WhenAll(addBatchTasks);

        var batchResponse = await graphClient.Batch.PostAsync(batchRequestContent);

        var responseTasks = requestIds.Select(id => batchResponse.GetResponseByIdAsync<T>(id)).ToList();

        while (responseTasks.Count > 0)
        {
            var completedTask = await Task.WhenAny(responseTasks);

            yield return await completedTask;

            responseTasks.Remove(completedTask);
        }
    }
}

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

public class AsyncEnumerableGraphDataService : IAsyncEnumerableGraphDataService
{
    public string? LastUrl { get; private set; } = null;

    private readonly GraphServiceClient _graphClient;

    private readonly RequestHeaders EventualConsistencyHeader = new() { { "ConsistencyLevel", "eventual" } };

    public AsyncEnumerableGraphDataService(GraphServiceClient graphClient)
    {
        _graphClient = graphClient;
    }

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
            });

        LastUrl = WebUtility.UrlDecode(requestInfo.URI.AbsoluteUri);
        return requestInfo.ToAsyncEnumerable<User, UserCollectionResponse>(_graphClient.RequestAdapter);
    }
}