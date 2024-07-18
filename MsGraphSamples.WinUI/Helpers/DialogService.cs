using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;

namespace MsGraphSamples.WinUI.Helpers;

public interface IDialogService
{
    XamlRoot? Root { get; set; }

    Task<ContentDialogResult> ShowAsync(string title, string? text, string closeButtonText = "Ok", string? primaryButtonText = null, string? secondaryButtonText = null);
}

public class DialogService() : IDialogService
{
    public XamlRoot? Root { get; set; }

    /// <summary>
    /// Shows a content dialog
    /// </summary>
    /// <param name="text">The text of the content dialog</param>
    /// <param name="title">The title of the content dialog</param>
    /// <param name="closeButtonText">The text of the close button</param>
    /// <param name="primaryButtonText">The text of the primary button (optional)</param>
    /// <param name="secondaryButtonText">The text of the secondary button (optional)</param>
    /// <returns>The ContentDialogResult</returns>
    public async Task<ContentDialogResult> ShowAsync(string title, string? text, string closeButtonText = "Ok", string? primaryButtonText = null, string? secondaryButtonText = null)
    {
        var dialog = new ContentDialog()
        {
            Title = title,
            Content = text,
            CloseButtonText = closeButtonText,
            PrimaryButtonText = primaryButtonText,
            SecondaryButtonText = secondaryButtonText,
            XamlRoot = Root
        };

        return await dialog.ShowAsync();
    }
}
