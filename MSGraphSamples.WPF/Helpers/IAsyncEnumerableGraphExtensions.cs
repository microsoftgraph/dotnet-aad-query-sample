using Microsoft.Graph;
using Microsoft.Graph.Applications;
using Microsoft.Graph.Models;
using Microsoft.Graph.Models.ExternalConnectors;
using Microsoft.Kiota.Abstractions;
using MsGraph_Samples.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static Microsoft.Graph.Applications.ApplicationsRequestBuilder;

namespace MsGraph_Samples.Helpers
{
    public static class IAsyncEnumerableGraphExtensions
    {
        /// <summary>
        /// Transform an ApplicationsCollection Request into an AsyncEnumerable to efficiently iterate through the collection in case there are several pages.
        /// </summary>
        /// <param name="request"></param>
        /// <returns></returns>
        public static async IAsyncEnumerable<T> ToAsyncEnumerable<T>(this RequestInformation requestInfo, IRequestAdapter requestAdapter) where T : Entity
        {
            do
            {
                var parsableCollection = await requestAdapter.SendAsync(requestInfo,
                    (parseNode) => new BaseCollectionPaginationCountResponse());

                var pageItems = parsableCollection.GetType()
                    .GetProperty("Value")
                    ?.GetValue(parsableCollection, null) as List<T>;

                if (pageItems == null)
                    break;

                foreach (T item in pageItems)
                {
                    yield return item;
                }
                
                var applicationsResponse = await request.GetAsync(rc => rc = requestConfiguration).ConfigureAwait(false);
                foreach (var item in applicationsResponse.Value)
                    yield return item;


                var authService = new AuthService();
                var graphServiceClient = authService.GraphClient;

                var usersResponse = await graphServiceClient
                    .Users
                    .GetAsync(requestConfiguration => { requestConfiguration.QueryParameters.Select = new string[] { "id", "createdDateTime" }; requestConfiguration.QueryParameters.Top = 1; });

                var userList = new List<User>();
                var pageIterator = PageIterator<User, UserCollectionResponse>.CreatePageIterator(
                    graphServiceClient,
                    usersResponse,
                    user =>
                    {
                        userList.Add(user);
                        return true;
                    });

                await pageIterator.IterateAsync();




                request. = 
            } while (request != null);
        }
    }
}
