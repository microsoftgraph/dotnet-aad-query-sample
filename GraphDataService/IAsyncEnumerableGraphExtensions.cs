using Microsoft.Graph.Models;
using Microsoft.Kiota.Abstractions;
using Microsoft.Kiota.Abstractions.Serialization;

namespace MsGraph_Samples.Helpers
{
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
            where TCollectionResponse : IParsable, IAdditionalDataHolder, new()
        {
            while(true)
            {
                var parsableCollection = await requestAdapter.SendAsync(requestInfo, parseNode => new TCollectionResponse()).ConfigureAwait(false);
                if (parsableCollection?.GetType().GetProperty("Value")?.GetValue(parsableCollection, null) is not List<TEntity> entities)
                {
                    // not a collection response
                    break;
                }

                foreach (var entity in entities)
                {
                    yield return entity;
                }

                if (parsableCollection.GetType().GetProperty("OdataNextLink")?.GetValue(parsableCollection) is not string nextLink)
                {
                    // no more pages
                    break;
                }

                requestInfo.URI = new Uri(nextLink);
            }
        }
    }
}

