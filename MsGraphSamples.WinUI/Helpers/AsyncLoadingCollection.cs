// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using System.Collections.ObjectModel;
using System.ComponentModel;
using Microsoft.UI.Xaml.Data;
using Windows.Foundation;

namespace MsGraphSamples.WinUI.Helpers;

/// <summary>
/// Initializes a new instance of the <see cref="AsyncLoadingCollection{T}"/> class.
/// </summary>
/// <param name="source">The source of items to load asynchronously.</param>
/// <param name="maxItemsPerPage">The maximum number of items to load per page.</param>
/// <param name="OnStartLoading">The action to invoke when loading starts.</param>
/// <param name="OnEndLoading">The action to invoke when loading ends.</param>
/// <param name="OnError">The action to invoke when an error occurs during loading.</param>
/// <param name="cancellationToken">The cancellation token to cancel the loading operation.</param>
public class AsyncLoadingCollection<T>(
    IAsyncEnumerable<T> source,
    uint defaultItemsPerPage = 25,
    CancellationToken cancellationToken = default)
    : ObservableCollection<T>, ISupportIncrementalLoading
{
    private IAsyncEnumerator<T>? _asyncEnumerator = source
        .GetAsyncEnumerator()
        .WithCancellation(cancellationToken);

    private readonly SemaphoreSlim _mutex = new(1, 1);
    public Action? OnStartLoading { get; set; }
    public Action? OnEndLoading { get; set; }
    public Action<Exception>? OnError { get; set; }

    public bool HasMoreItems => _asyncEnumerator != null;

    private bool _isLoading;
    /// <summary>
    /// Gets a value indicating whether new items are being loaded.
    /// </summary>
    public bool IsLoading
    {
        get => _isLoading;

        private set
        {
            if (value == _isLoading)
                return;

            _isLoading = value;
            OnPropertyChanged(new PropertyChangedEventArgs(nameof(IsLoading)));

            if (_isLoading)
                OnStartLoading?.Invoke();
            else
                OnEndLoading?.Invoke();
        }
    }
    public IAsyncOperation<LoadMoreItemsResult> LoadMoreItemsAsync(uint count) =>
        LoadMoreItemsAsyncInternal(count)
        .AsAsyncOperation();

    private async Task<LoadMoreItemsResult> LoadMoreItemsAsyncInternal(uint count)
    {
        await _mutex.WaitAsync(cancellationToken);

        uint itemsLoaded = 0;
        IsLoading = true;

        try
        {
            while (itemsLoaded < count && HasMoreItems)
            {
                if (await _asyncEnumerator!.MoveNextAsync().ConfigureAwait(false))
                {
                    Add(_asyncEnumerator!.Current);
                    itemsLoaded++;
                }
                else
                {
                    // Dispose the enumerator when we're done
                    await _asyncEnumerator!.DisposeAsync();
                    _asyncEnumerator = null;
                }
            }
        }
        catch (OperationCanceledException)
        {
            // The operation has been canceled using the Cancellation Token.
            await _asyncEnumerator!.DisposeAsync();
            _asyncEnumerator = null;
        }
        catch (Exception ex)
        {
            OnError?.Invoke(ex);
            await _asyncEnumerator!.DisposeAsync();
            _asyncEnumerator = null;
            throw;
        }
        finally
        {
            IsLoading = false;
            _mutex.Release();
        }

        return new LoadMoreItemsResult(itemsLoaded);
    }

    /// <summary>
    /// Clears the collection and triggers/forces a reload of the first page
    /// </summary>
    /// <returns>
    /// An object of the <see cref="LoadMoreItemsAsync(uint)"/> that specifies how many items have been actually retrieved.
    /// </returns>
    public async Task<LoadMoreItemsResult> RefreshAsync()
    {
        await _mutex.WaitAsync(cancellationToken);

        Clear();
        _asyncEnumerator = source
            .GetAsyncEnumerator()
            .WithCancellation(cancellationToken);

        _mutex.Release();

        return await LoadMoreItemsAsync(defaultItemsPerPage);
    }
}