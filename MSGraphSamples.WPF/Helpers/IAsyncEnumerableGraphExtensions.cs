using Microsoft.Graph.Models;
using Microsoft.Kiota.Abstractions;
using Microsoft.Kiota.Abstractions.Serialization;

namespace MsGraph_Samples.Helpers
{
    public static class IAsyncEnumerableGraphExtensions
    {
        ///// <summary>
        ///// Transform an ApplicationCollectionResponse into an AsyncEnumerable to efficiently iterate through the collection in case there are several pages.
        ///// </summary>
        ///// <param name="requestInfo"></param>
        ///// <param name="requestAdapter"></param>
        ///// <returns>IAsyncEnumerable<Application></returns>
        //public static async IAsyncEnumerable<Microsoft.Graph.Models.Application> ToAsyncEnumerable<Application>(this RequestInformation requestInfo, IRequestAdapter requestAdapter)
        //{
        //    var apps = await requestAdapter.SendAsync(requestInfo, (parseNode) => new ApplicationCollectionResponse()).ConfigureAwait(false);
        //    foreach (var item in apps.Value)
        //    {
        //        yield return item;
        //    }

        //    while (apps.OdataNextLink != null)
        //    {
        //        requestInfo.URI = new Uri(apps.OdataNextLink);
        //        apps = await requestAdapter.SendAsync(requestInfo, (parseNode) => new ApplicationCollectionResponse()).ConfigureAwait(false);
        //        foreach (var item in apps.Value)
        //        {
        //            yield return item;
        //        }
        //    }
        //}

        /// <summary>
        /// Transform an UserCollectionResponse into an AsyncEnumerable to efficiently iterate through the collection in case there are several pages.
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
                if (parsableCollection.GetType().GetProperty("Value")?.GetValue(parsableCollection, null) is not List<TEntity> entities)
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


        //parsableCollection.AdditionalData.TryGetValue("@odata.nextLink", out var nextLinkValue);

        //await requestAdapter.SendAsync(requestInfo, (parseNode) => new BaseCollectionPaginationCountResponse());            
        //var element = (JsonElement)parsableCollection.AdditionalData["value"];
        //var users = element.Deserialize<User[]>();


        //var applicationsResponse = await request.GetAsync(rc => rc = requestConfiguration).ConfigureAwait(false);
        //foreach (var item in applicationsResponse.Value)
        //    yield return item;


        //var authService = new AuthService();
        //var graphServiceClient = authService.GraphClient;

        //var usersResponse = await graphServiceClient
        //    .Users
        //    .GetAsync(requestConfiguration => { requestConfiguration.QueryParameters.Select = new string[] { "id", "createdDateTime" }; requestConfiguration.QueryParameters.Top = 1; });

        //var userList = new List<User>();
        //var pageIterator = PageIterator<User, UserCollectionResponse>.CreatePageIterator(
        //    graphServiceClient,
        //    usersResponse,
        //    user =>
        //    {
        //        userList.Add(user);
        //        return true;
        //    });

        //await pageIterator.IterateAsync();
    }
}

