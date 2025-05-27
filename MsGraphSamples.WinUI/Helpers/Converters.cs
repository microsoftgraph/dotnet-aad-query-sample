// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using Microsoft.UI.Xaml.Data;
using System.Collections;

namespace MsGraphSamples.WinUI.Converters;

public partial class AdditionalDataConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, string language)
    {
        var additionalData = (IDictionary<string, object>)value;
        additionalData.TryGetValue((string)parameter, out var extensionValue);

        return extensionValue?.ToString() ?? string.Empty;
    }

    public object ConvertBack(object value, Type targetType, object parameter, string language) => throw new NotImplementedException();
}
public class CollectionToCommaSeparatedStringConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, string language)
    {
        if (value is IEnumerable<string> stringEnumerable)
        {
            return string.Join(", ", stringEnumerable);
        }
        if (value is IEnumerable enumerable && value is not string)
        {
            var items = new List<string>();
            foreach (var item in enumerable)
            {
                if (item != null)
                    items.Add(item.ToString());
            }
            return string.Join(", ", items);
        }
        return value?.ToString() ?? string.Empty;
    }

    public object ConvertBack(object value, Type targetType, object parameter, string language)
    {
        throw new NotImplementedException();
    }
}