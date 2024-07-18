// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using System.Collections.ObjectModel;
using Microsoft.UI.Xaml.Data;
using Windows.Foundation;

namespace MsGraphSamples.WinUI.Helpers;

public class AsyncLoadingCollection<T>(IAsyncEnumerable<T> source, uint itemsPerPage = 25) : ObservableCollection<T>, ISupportIncrementalLoading
{
    private IAsyncEnumerator<T>? _asyncEnumerator = source.GetAsyncEnumerator();
    private readonly SemaphoreSlim _mutex = new(1, 1);

    public bool HasMoreItems => _asyncEnumerator != null;

    public IAsyncOperation<LoadMoreItemsResult> LoadMoreItemsAsync(uint count = 0) =>
        LoadMoreItemsAsync(count == 0 ? itemsPerPage : count, default)
        .AsAsyncOperation();

    private async Task<LoadMoreItemsResult> LoadMoreItemsAsync(uint count, CancellationToken cancellationToken)
    {
        await _mutex.WaitAsync(cancellationToken);

        if (cancellationToken.IsCancellationRequested || !HasMoreItems)
            return new LoadMoreItemsResult(0);

        uint itemsLoaded = 0;
        var itemsToLoad = Math.Min(itemsPerPage, count);

        try
        {
            while (itemsLoaded < itemsToLoad)
            {
                if (await _asyncEnumerator!.MoveNextAsync(cancellationToken).ConfigureAwait(false))
                {
                    Add(_asyncEnumerator!.Current);
                    itemsLoaded++;
                }
                else
                {
                    // Dispose the enumerator when we're done
                    await _asyncEnumerator!.DisposeAsync();
                    _asyncEnumerator = null;
                    break;
                }
            }
        }
        catch (OperationCanceledException)
        {
            // The operation has been canceled using the Cancellation Token.
            await _asyncEnumerator!.DisposeAsync();
            _asyncEnumerator = null;
        }
        catch (Exception)
        {
            await _asyncEnumerator!.DisposeAsync();
            _asyncEnumerator = null;
            throw;
        }
        finally
        {
            _mutex.Release();
        }

        return new LoadMoreItemsResult(itemsLoaded);
    }
}