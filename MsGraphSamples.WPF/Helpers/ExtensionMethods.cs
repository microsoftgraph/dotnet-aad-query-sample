// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using System.Collections.Specialized;
using System.Web;

namespace MsGraph_Samples;

public static class ExtensionMethods
{
    public static bool In(this string s, params string[] items) => items.Any(i => i.Trim().Equals(s, StringComparison.InvariantCultureIgnoreCase));

    public static int NthIndexOf(this string input, char value, int nth, int startIndex = 0)
    {
        if (nth <= 0)
            throw new ArgumentException("Input must be greater than 0", nameof(nth));
        if (nth == 1)
            return input.IndexOf(value, startIndex);

        return input.NthIndexOf(value, --nth, input.IndexOf(value, startIndex) + 1);
    }

    public static string? Decode(this Uri? uri) => HttpUtility.UrlDecode(uri?.ToString());
    public static NameValueCollection ParseQueryString(this Uri uri) => HttpUtility.ParseQueryString(uri.Query);

    /// <summary>
    /// Awaits a task without blocking the main thread. (From PRISM framework)
    /// </summary>
    /// <remarks>Primarily used to replace async void scenarios such as ctor's and ICommands.</remarks>
    /// <param name="task">The task to be awaited</param>
    /// <param name="completedCallback">The action to perform when the task is complete.</param>
    /// <param name="errorCallback">The action to perform when an error occurs executing the task.</param>
    /// <param name="configureAwait">Configures an awaiter used to await this task</param>
    public async static void Await(this Task task, Action? completedCallback = null, Action<Exception>? errorCallback = null, bool configureAwait = false)
    {
        try
        {
            await task.ConfigureAwait(configureAwait);
            completedCallback?.Invoke();
        }
        catch (Exception ex)
        {
            errorCallback?.Invoke(ex);
        }
    }

    /// <summary>
    /// Awaits a task without blocking the main thread. (From PRISM framework)
    /// </summary>
    /// <remarks>Primarily used to replace async void scenarios such as ctor's and ICommands.</remarks>
    /// <param name="task">The task to be awaited</param>
    /// <param name="completedCallback">The action to perform when the task is complete.</param>
    /// <param name="errorCallback">The action to perform when an error occurs executing the task.</param>
    /// <param name="configureAwait">Configures an awaiter used to await this task</param>
    public async static void Await<T>(this Task<T> task, Action<T>? completedCallback = null, Action<Exception>? errorCallback = null, bool configureAwait = false)
    {
        try
        {
            T result = await task.ConfigureAwait(configureAwait);
            completedCallback?.Invoke(result);
        }
        catch (Exception ex)
        {
            errorCallback?.Invoke(ex);
        }
    }
}