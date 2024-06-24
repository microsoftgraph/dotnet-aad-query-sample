// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using Microsoft.Graph.Models;
using Microsoft.Kiota.Abstractions.Store;
using System.Globalization;
using System.Windows.Data;

namespace MsGraphSamples.WPF.Converters;

public class AdditionalDataConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        var additionalData = (IDictionary<string, object>)value;
        additionalData.TryGetValue((string)parameter, out var extensionValue);

        return extensionValue?.ToString() ?? string.Empty;
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture) => throw new NotImplementedException();
}

public class DirectoryObjectsCountConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        var directoryObjects = (BaseCollectionPaginationCountResponse?)value;

        if (directoryObjects == null)
            return string.Empty;

        var directoryObjectCollection = directoryObjects.BackingStore?.Get<IEnumerable<DirectoryObject>?>("value");
        return $"{directoryObjectCollection?.Count() ?? 0} / {directoryObjects.OdataCount}";
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture) => throw new NotImplementedException();
}