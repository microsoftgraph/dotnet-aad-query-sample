using Microsoft.Graph.Models;
using Microsoft.Graph;
using Microsoft.Kiota.Abstractions.Serialization;
using Microsoft.Kiota.Abstractions;

namespace MsGraphSamples.Services;

public static class GraphExtensions
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

        while (collectionResponse != null)
        {
            var entities = collectionResponse.GetValue<TEntity>() ?? [];
            foreach (var entity in entities)
            {
                yield return entity;
            }

            collectionResponse = await collectionResponse.GetNextPageAsync(requestAdapter);
        }
    }

    public static List<TEntity>? GetValue<TEntity>(this BaseCollectionPaginationCountResponse collectionResponse) where TEntity : Entity
    {
        return collectionResponse.BackingStore.Get<List<TEntity>>("value");
    }

    public static async Task<TCollectionResponse?> GetNextPageAsync<TCollectionResponse>(this TCollectionResponse? collectionResponse, IRequestAdapter requestAdapter)
        where TCollectionResponse : BaseCollectionPaginationCountResponse, new()
    {
        if (collectionResponse?.OdataNextLink == null)
            return null;

        var nextPageRequestInformation = new RequestInformation
        {
            HttpMethod = Method.GET,
            UrlTemplate = collectionResponse.OdataNextLink,
        };
        var previousCount = collectionResponse.OdataCount;

        var nextPage = await requestAdapter
            .SendAsync(nextPageRequestInformation, parseNode => new TCollectionResponse())
            .ConfigureAwait(false);

        // fix count property not present in pages other than the first one
        if (nextPage != null)
            nextPage.OdataCount = previousCount;

        return nextPage;
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
