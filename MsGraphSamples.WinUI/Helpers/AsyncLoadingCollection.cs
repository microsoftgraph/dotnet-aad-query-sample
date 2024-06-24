// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using System.Collections.ObjectModel;
using Microsoft.UI.Xaml.Data;
using Windows.Foundation;

namespace MsGraphSamples.WinUI.Helpers;

public class AsyncLoadingCollection<T> : ObservableCollection<T>, ISupportIncrementalLoading
{
    private readonly uint _itemsPerPage;
    private IAsyncEnumerator<T>? _asyncEnumerator;
    private readonly SemaphoreSlim _mutex = new(1, 1);

    public bool HasMoreItems => _asyncEnumerator != null;

    public AsyncLoadingCollection(IAsyncEnumerable<T> source, uint itemsPerPage = 25)
    {
        ArgumentNullException.ThrowIfNull(nameof(source));

        _asyncEnumerator = source.GetAsyncEnumerator();
        _itemsPerPage = itemsPerPage;
    }

    public IAsyncOperation<LoadMoreItemsResult> LoadMoreItemsAsync(uint count = default) =>
        LoadMoreItemsAsync(count == default ? _itemsPerPage : count, new CancellationToken(false))
        .AsAsyncOperation();

    private async Task<LoadMoreItemsResult> LoadMoreItemsAsync(uint count, CancellationToken cancellationToken)
    {
        await _mutex.WaitAsync(cancellationToken);

        if (cancellationToken.IsCancellationRequested || !HasMoreItems)
        {
            return new LoadMoreItemsResult(0);
        }

        uint itemsLoaded = 0;
        var itemsToLoad = Math.Min(_itemsPerPage, count);

        try
        {
            while (itemsLoaded < itemsToLoad)
            {
                if (await _asyncEnumerator!.MoveNextAsync())
                {
                    Add(_asyncEnumerator.Current);
                    itemsLoaded++;
                }
                else
                {
                    // Dispose the enumerator when we're done
                    await _asyncEnumerator.DisposeAsync();
                    _asyncEnumerator = null;
                    break;
                }

                if (cancellationToken.IsCancellationRequested)
                {
                    break;
                }
            }
        }
        catch (OperationCanceledException)
        {
            // The operation has been canceled using the Cancellation Token.
        }
        catch (Exception)
        {
            throw;
        }
        finally
        {
            _mutex.Release();
        }

        return new LoadMoreItemsResult(itemsLoaded);
    }
}
