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

    /// <summary>
    /// Transform a generic BaseCollectionPaginationCountResponse into an AsyncEnumerable to efficiently iterate through the collection in case there are several pages.
    /// </summary>
    /// <typeparam name="TEntity">Microsoft Graph Entity of the CollectionResponse</typeparam>
    /// <typeparam name="TCollectionResponse">Specialized BaseCollectionPaginationCountResponse</typeparam>
    /// <param name="collectionResponse">The CollectionResponse to convert to IAsyncEnumerable</param>
    /// <param name="requestAdapter">The IRequestAdapter from GraphServiceClient used to make requests</param>
    /// <returns></returns>
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
