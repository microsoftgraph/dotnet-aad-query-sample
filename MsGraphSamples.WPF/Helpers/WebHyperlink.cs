﻿using System.Diagnostics;
using System.Windows;
using System.Windows.Documents;
using System.Windows.Navigation;

namespace MsGraph_Samples.Helpers;

class HyperlinkExtensions
{
    public static bool GetIsWeb(DependencyObject obj)
    {
        return (bool)obj.GetValue(IsWebProperty);
    }

    public static void SetIsWeb(DependencyObject obj, bool value)
    {
        obj.SetValue(IsWebProperty, value);
    }

    public static readonly DependencyProperty IsWebProperty = DependencyProperty.RegisterAttached(
        "IsWeb",
        typeof(bool),
        typeof(HyperlinkExtensions),
        new UIPropertyMetadata(false, OnIsExternalChanged));

    private static void OnIsExternalChanged(object sender, DependencyPropertyChangedEventArgs args)
    {
        var hyperlink = (Hyperlink)sender;

        if ((bool)args.NewValue)
            hyperlink.RequestNavigate += Hyperlink_RequestNavigate;
        else
            hyperlink.RequestNavigate -= Hyperlink_RequestNavigate;
    }
    private static void Hyperlink_RequestNavigate(object sender, RequestNavigateEventArgs e)
    {
        var psi = new ProcessStartInfo(e.Uri.AbsoluteUri) { UseShellExecute = true };
        Process.Start(psi);
        e.Handled = true;
    }
}