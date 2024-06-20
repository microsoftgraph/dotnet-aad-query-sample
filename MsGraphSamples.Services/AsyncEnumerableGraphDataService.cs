// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using System.Net;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Kiota.Abstractions;

namespace MsGraphSamples.Services;

public static class IAsyncEnumerableGraphExtensions
{
    /// <summary>
    /// Transform a generic RequestInformation into an AsyncEnumerable to efficiently iterate through the collection in case there are several pages.
    /// </summary>
    /// <param name="requestInfo"></param>
    /// <param name="requestAdapter"></param>
    /// <param name="countAction"></param>
    /// <returns>IAsyncEnumerable<Entity></returns>
    public static IAsyncEnumerable<TEntity> ToAsyncEnumerable<TEntity, TCollectionResponse>(this RequestInformation requestInfo, IRequestAdapter requestAdapter, Action<long?>? countAction = null)
        where TEntity : Entity
        where TCollectionResponse : BaseCollectionPaginationCountResponse, new()
    {
        return requestAdapter
            .SendAsync(requestInfo, parseNode => new TCollectionResponse())
            .ToAsyncEnumerable<TEntity, TCollectionResponse>(requestAdapter, countAction);
    }

    /// <summary>
    /// Transform a Task<BaseCollectionPaginationCountResponse> into an AsyncEnumerable to efficiently iterate through the collection in case there are several pages.
    /// </summary>
    /// <param name="requestInfo"></param>
    /// <param name="requestAdapter"></param>
    /// <param name="countAction"></param>
    /// <returns>IAsyncEnumerable<Entity></returns>
    public static async IAsyncEnumerable<TEntity> ToAsyncEnumerable<TEntity, TCollectionResponse>(this Task<TCollectionResponse?> collectionResponseTask, IRequestAdapter requestAdapter, Action<long?>? countAction = null)
        where TEntity : Entity
        where TCollectionResponse : BaseCollectionPaginationCountResponse, new()
    {
        var collectionResponse = await collectionResponseTask.ConfigureAwait(false);
        await foreach (var item in collectionResponse.ToAsyncEnumerable<TEntity, TCollectionResponse>(requestAdapter, countAction))
        {
            yield return item;
        }
    }

    /// <summary>
    /// Transform a BaseCollectionPaginationCountResponse into an AsyncEnumerable to efficiently iterate through the collection in case there are several pages.
    /// </summary>
    /// <typeparam name="TEntity"></typeparam>
    /// <typeparam name="TCollectionResponse"></typeparam>
    /// <param name="collectionResponse"></param>
    /// <param name="requestAdapter"></param>
    /// <param name="countAction"></param>
    /// <returns></returns>
    public static async IAsyncEnumerable<TEntity> ToAsyncEnumerable<TEntity, TCollectionResponse>(this TCollectionResponse? collectionResponse, IRequestAdapter requestAdapter, Action<long?>? countAction = null)
    where TEntity : Entity
    where TCollectionResponse : BaseCollectionPaginationCountResponse, new()
    {
        countAction?.Invoke(collectionResponse?.OdataCount);

        while (true)
        {
            var entities = collectionResponse?.BackingStore.Get<List<TEntity>>("value") ?? Enumerable.Empty<TEntity>();
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


    public static async IAsyncEnumerable<T> Batch<T>(this GraphServiceClient graphClient, params RequestInformation[] requests)
        where T : IParsable, new()
    {
        BatchRequestContentCollection batchRequestContent = new(graphClient);

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
}

public interface IAsyncEnumerableGraphDataService
{
    string? LastUrl
    {
        get;
    }
    public long? LastCount
    {
        get;
    }

    Task<User?> GetUserAsync(string[] select, string? id = null);
    IAsyncEnumerable<User> GetUsersInBatch(string[] select, ushort? top = null);
    IAsyncEnumerable<Application> GetApplications(string[] select, string? filter = null, string[]? orderBy = null, string? search = null, ushort? top = null);
    IAsyncEnumerable<ServicePrincipal> GetServicePrincipals(string[] splittedSelect, string? filter, string[]? splittedOrderBy, string? search, ushort? top = null);
    IAsyncEnumerable<Device> GetDevices(string[] select, string? filter = null, string[]? orderBy = null, string? search = null, ushort? top = null);
    IAsyncEnumerable<Group> GetGroups(string[] select, string? filter = null, string[]? orderBy = null, string? search = null, ushort? top = null);
    IAsyncEnumerable<User> GetUsers(string[] select, string? filter = null, string[]? orderBy = null, string? search = null, ushort? top = null);
    IAsyncEnumerable<Group> GetTransitiveMemberOfAsGroups(string id, string[] select, ushort? top = null);
    IAsyncEnumerable<User> GetTransitiveMembersAsUsers(string id, string[] select, ushort? top = null);
    IAsyncEnumerable<User> GetAppOwnersAsUsers(string id, string[] select, ushort? top = null);
    IAsyncEnumerable<User> GetSPOwnersAsUsers(string id, string[] select, ushort? top = null);
    IAsyncEnumerable<User> GetRegisteredOwnersAsUsers(string id, string[] select, ushort? top = null);
}

public class AsyncEnumerableGraphDataService(GraphServiceClient graphClient) : IAsyncEnumerableGraphDataService
{
    public string? LastUrl { get; private set; } = null;
    public long? LastCount { get; private set; } = null;
    private void SetCount(long? count) => LastCount = count;

    private readonly GraphServiceClient _graphClient;

    private readonly RequestHeaders EventualConsistencyHeader = new() { { "ConsistencyLevel", "eventual" } };

    public string? LastUrl { get; private set; } = null;

    public Task<User?> GetUserAsync(string[] select, string? id = null)
    {
        return id == null
            ? _graphClient.Me.GetAsync(rc => rc.QueryParameters.Select = select)
            : _graphClient.Users[id].GetAsync(rc => rc.QueryParameters.Select = select);
    }

    public IAsyncEnumerable<User> GetUsersInBatch(string[] select, ushort? top = null)
    {
        return _graphClient.Batch<User, UserCollectionResponse>(
            _graphClient.Users.ToGetRequestInformation(rc =>
            {
                rc.Headers = EventualConsistencyHeader;
                rc.QueryParameters.Count = true;
                rc.QueryParameters.Select = select;
                rc.QueryParameters.Filter = "startsWith(displayName, 'x')";
                rc.QueryParameters.Top = top;
            }),
            _graphClient.Users.ToGetRequestInformation(rc =>
            {
                rc.Headers = EventualConsistencyHeader;
                rc.QueryParameters.Count = true;
                rc.QueryParameters.Select = select;
                rc.QueryParameters.Filter = "startsWith(displayName, 'y')";
                rc.QueryParameters.Top = top;
            }),
            _graphClient.Users.ToGetRequestInformation(rc =>
            {
                rc.Headers = EventualConsistencyHeader;
                rc.QueryParameters.Count = true;
                rc.QueryParameters.Select = select;
                rc.QueryParameters.Filter = "startsWith(displayName, 'z')";
                rc.QueryParameters.Top = top;
            }));
    }

    public IAsyncEnumerable<Application> GetApplications(string[] select, string? filter = null, string[]? orderBy = null, string? search = null, ushort? top = null)
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
                rc.QueryParameters.Top = top;
            });

        LastUrl = WebUtility.UrlDecode(requestInfo.URI.AbsoluteUri);
        return requestInfo.ToAsyncEnumerable<Application, ApplicationCollectionResponse>(_graphClient.RequestAdapter, SetCount);
    }

    public IAsyncEnumerable<ServicePrincipal> GetServicePrincipals(string[] splittedSelect, string? filter, string[]? splittedOrderBy, string? search, ushort? top = null)
    {
        var requestInfo = _graphClient.ServicePrincipals
            .ToGetRequestInformation(rc =>
            {
                rc.Headers = EventualConsistencyHeader;
                rc.QueryParameters.Count = true;
                rc.QueryParameters.Select = splittedSelect;
                rc.QueryParameters.Filter = filter;
                rc.QueryParameters.Orderby = splittedOrderBy;
                rc.QueryParameters.Search = search;
                rc.QueryParameters.Top = top;
            });

        LastUrl = WebUtility.UrlDecode(requestInfo.URI.AbsoluteUri);
        return requestInfo.ToAsyncEnumerable<ServicePrincipal, ServicePrincipalCollectionResponse>(_graphClient.RequestAdapter, SetCount);
    }

    public IAsyncEnumerable<Device> GetDevices(string[] select, string? filter = null, string[]? orderBy = null, string? search = null, ushort? top = null)
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
                rc.QueryParameters.Top = top;
            });

        LastUrl = WebUtility.UrlDecode(requestInfo.URI.AbsoluteUri);
        return requestInfo.ToAsyncEnumerable<Device, DeviceCollectionResponse>(_graphClient.RequestAdapter, SetCount);
    }

    public IAsyncEnumerable<Group> GetGroups(string[] select, string? filter = null, string[]? orderBy = null, string? search = null, ushort? top = null)
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
                rc.QueryParameters.Top = top;
            });

        LastUrl = WebUtility.UrlDecode(requestInfo.URI.AbsoluteUri);
        return requestInfo.ToAsyncEnumerable<Group, GroupCollectionResponse>(_graphClient.RequestAdapter, SetCount);
    }

    public IAsyncEnumerable<User> GetUsers(string[] select, string? filter = null, string[]? orderBy = null, string? search = null, ushort? top = null)
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
            rc.QueryParameters.Top = top;
        });

        LastUrl = WebUtility.UrlDecode(requestInfo.URI.AbsoluteUri);
        return requestInfo.ToAsyncEnumerable<User, UserCollectionResponse>(_graphClient.RequestAdapter, SetCount);
    }

    public IAsyncEnumerable<Group> GetTransitiveMemberOfAsGroups(string id, string[] select, ushort? top = null)
    {
        var requestInfo = _graphClient.Users[id]
            .TransitiveMemberOf.GraphGroup
            .ToGetRequestInformation(rc =>
            {
                rc.Headers = EventualConsistencyHeader;
                rc.QueryParameters.Count = true;
                rc.QueryParameters.Select = select;
                rc.QueryParameters.Top = top;
            });

        LastUrl = WebUtility.UrlDecode(requestInfo.URI.AbsoluteUri);
        return requestInfo.ToAsyncEnumerable<Group, GroupCollectionResponse>(_graphClient.RequestAdapter, SetCount);
    }

    public IAsyncEnumerable<User> GetTransitiveMembersAsUsers(string id, string[] select, ushort? top = null)
    {
        var requestInfo = _graphClient.Groups[id]
            .TransitiveMembers.GraphUser
            .ToGetRequestInformation(rc =>
            {
                rc.Headers = EventualConsistencyHeader;
                rc.QueryParameters.Count = true;
                rc.QueryParameters.Select = select;
                rc.QueryParameters.Top = top;
            });

        LastUrl = WebUtility.UrlDecode(requestInfo.URI.AbsoluteUri);
        return requestInfo.ToAsyncEnumerable<User, UserCollectionResponse>(_graphClient.RequestAdapter, SetCount);
    }

    public IAsyncEnumerable<User> GetAppOwnersAsUsers(string id, string[] select, ushort? pageSize = null)
    {
        var requestInfo = _graphClient.Applications[id]
            .Owners.GraphUser
            .ToGetRequestInformation(rc =>
            {
                rc.Headers = EventualConsistencyHeader;
                rc.QueryParameters.Count = true;
                rc.QueryParameters.Select = select;
                rc.QueryParameters.Top = pageSize;
            });

        LastUrl = WebUtility.UrlDecode(requestInfo.URI.AbsoluteUri);
        return requestInfo.ToAsyncEnumerable<User, UserCollectionResponse>(_graphClient.RequestAdapter, SetCount);
    }

    public IAsyncEnumerable<User> GetSPOwnersAsUsers(string id, string[] select, ushort? top = null)
    {
        var requestInfo = _graphClient.ServicePrincipals[id]
            .Owners.GraphUser
            .ToGetRequestInformation(rc =>
            {
                
                rc.Headers = EventualConsistencyHeader;
                rc.QueryParameters.Count = true;
                rc.QueryParameters.Select = select;
                rc.QueryParameters.Top = top;
            });

        LastUrl = WebUtility.UrlDecode(requestInfo.URI.AbsoluteUri);
        return requestInfo.ToAsyncEnumerable<User, UserCollectionResponse>(_graphClient.RequestAdapter, SetCount);
    }

    public IAsyncEnumerable<User> GetRegisteredOwnersAsUsers(string id, string[] select, ushort? top = null)
    {
        var requestInfo = _graphClient.Devices[id]
            .RegisteredOwners.GraphUser
            .ToGetRequestInformation(rc =>
            {

                rc.Headers = EventualConsistencyHeader;
                rc.QueryParameters.Count = true;
                rc.QueryParameters.Select = select;
                rc.QueryParameters.Top = top;
            });

        LastUrl = WebUtility.UrlDecode(requestInfo.URI.AbsoluteUri);
        return requestInfo.ToAsyncEnumerable<User, UserCollectionResponse>(_graphClient.RequestAdapter, SetCount);
    }

}