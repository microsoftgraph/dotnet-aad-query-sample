using Microsoft.Graph.Models;
using Microsoft.Graph;
using Microsoft.Kiota.Abstractions.Serialization;
using Microsoft.Kiota.Abstractions;

namespace MsGraphSamples.Services;

public static class AsyncEnumerableGraphExtensions
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
    /// Transform a generic BaseCollectionPaginationCountResponse into an AsyncEnumerable to efficiently iterate through the collection in case there are several pages.
    /// </summary>
    /// <typeparam name="TEntity">Microsoft Graph Entity of the CollectionResponse</typeparam>
    /// <typeparam name="TCollectionResponse">Specialized BaseCollectionPaginationCountResponse</typeparam>
    /// <param name="collectionResponse">The CollectionResponse to convert to IAsyncEnumerable</param>
    /// <param name="requestAdapter">The IRequestAdapter from GraphServiceClient used to make requests</param>
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
}
