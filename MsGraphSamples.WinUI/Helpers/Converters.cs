// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using Microsoft.UI.Xaml.Data;

namespace MsGraphSamples.WinUI.Converters;

public class AdditionalDataConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, string language)
    {
        var additionalData = (IDictionary<string, object>)value;
        additionalData.TryGetValue((string)parameter, out var extensionValue);

        return extensionValue?.ToString() ?? string.Empty;
    }

    public object ConvertBack(object value, Type targetType, object parameter, string language) => throw new NotImplementedException();
}