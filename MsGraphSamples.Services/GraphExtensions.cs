using Microsoft.Graph.Models;
using Microsoft.Graph;
using Microsoft.Kiota.Abstractions.Serialization;
using Microsoft.Kiota.Abstractions;
using Microsoft.Graph.Models.ODataErrors;
using System.Runtime.CompilerServices;

namespace MsGraphSamples.Services;

public static class GraphExtensions
{
    private static readonly Dictionary<string, ParsableFactory<IParsable>> ErrorMappings = new() { { "XXX", ODataError.CreateFromDiscriminatorValue } };

    /// <summary>
    /// Transform a generic RequestInformation into an AsyncEnumerable to efficiently iterate through the collection in case there are several pages.
    /// </summary>
    /// <param name="requestInfo"></param>
    /// <param name="requestAdapter"></param>
    /// <param name="countAction"></param>
    /// <returns>IAsyncEnumerable<Entity></returns>
    public static async IAsyncEnumerable<TEntity> ToAsyncEnumerable<TEntity, TCollectionResponse>(this RequestInformation requestInfo, IRequestAdapter requestAdapter, Action<long?>? countAction = null, [EnumeratorCancellation] CancellationToken cancellationToken = default)
        where TEntity : Entity
        where TCollectionResponse : BaseCollectionPaginationCountResponse, new()
    {
        // Send the asynchronous request and get the response
        var collectionResponse = await requestAdapter
            .SendAsync(requestInfo, parseNode => new TCollectionResponse(), ErrorMappings, cancellationToken);

        // Iterate through the collection response asynchronously
        await foreach (var item in collectionResponse.ToAsyncEnumerable<TEntity, TCollectionResponse>(requestAdapter, countAction, cancellationToken))
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
    public static async IAsyncEnumerable<TEntity> ToAsyncEnumerable<TEntity, TCollectionResponse>(this TCollectionResponse? collectionResponse, IRequestAdapter requestAdapter, Action<long?>? countAction = null, [EnumeratorCancellation] CancellationToken cancellationToken = default)
    where TEntity : Entity
    where TCollectionResponse : BaseCollectionPaginationCountResponse, new()
    {
        countAction?.Invoke(collectionResponse?.OdataCount);

        while (collectionResponse != null)
        {
            var entities = collectionResponse.GetValue<TEntity>();
            foreach (var entity in entities)
            {
                yield return entity;
            }

            collectionResponse = await collectionResponse
                .GetNextPageAsync(requestAdapter, cancellationToken)
                .ConfigureAwait(false);
        }
    }

    public static IList<TEntity> GetValue<TEntity>(this BaseCollectionPaginationCountResponse? collectionResponse) where TEntity : Entity
    {
        return collectionResponse?.BackingStore.Get<IList<TEntity>>("value") ?? [];
    }

    public static async Task<TCollectionResponse?> GetNextPageAsync<TCollectionResponse>(this TCollectionResponse? collectionResponse, IRequestAdapter requestAdapter, CancellationToken cancellationToken = default)
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
            .SendAsync(nextPageRequestInformation, parseNode => new TCollectionResponse(), ErrorMappings, cancellationToken)
            .ConfigureAwait(false);

        // fix count property not present in pages other than the first one
        if (nextPage != null)
            nextPage.OdataCount = previousCount;

        return nextPage;
    }


    public static async IAsyncEnumerable<TEntity> Batch<TEntity, TCollectionResponse>(this GraphServiceClient graphClient, [EnumeratorCancellation] CancellationToken cancellationToken = default, params RequestInformation[] requests)
        where TEntity : Entity
        where TCollectionResponse : BaseCollectionPaginationCountResponse, new()
    {
        await foreach (var response in graphClient.Batch<TCollectionResponse>(cancellationToken, requests))
        {
            await foreach (var entity in response
                .ToAsyncEnumerable<TEntity, TCollectionResponse>(graphClient.RequestAdapter)
                .WithCancellation(cancellationToken)
                .ConfigureAwait(false))
            {
                yield return entity;
            }
        }
    }

    public static async IAsyncEnumerable<T> Batch<T>(
        this GraphServiceClient graphClient,
        [EnumeratorCancellation] CancellationToken cancellationToken = default,
        params RequestInformation[] requests)
        where T : IParsable, new()
    {
        BatchRequestContentCollection batchRequestContent = new(graphClient);

        var addBatchTasks = requests.Select(x => batchRequestContent.AddBatchRequestStepAsync(x));
        var requestIds = await Task.WhenAll(addBatchTasks);

        var batchResponse = await graphClient.Batch.PostAsync(batchRequestContent, cancellationToken, ErrorMappings);

        var responseTasks = requestIds.Select(id => batchResponse.GetResponseByIdAsync<T>(id)).ToList();

        // return first response as soon as it's available
        while (responseTasks.Count > 0)
        {
            var completedTask = await Task.WhenAny(responseTasks);
            yield return await completedTask;
            responseTasks.Remove(completedTask);
        }
    }
}