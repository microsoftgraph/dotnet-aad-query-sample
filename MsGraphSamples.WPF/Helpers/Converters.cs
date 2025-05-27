// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using Microsoft.Graph.Models;
using MsGraphSamples.Services;
using System.Globalization;
using System.Windows.Data;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

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

        var directoryObjectCollection = directoryObjects.BackingStore.Get<IEnumerable<DirectoryObject>>("value") ?? [];
        return $"{directoryObjectCollection.Count()} / {directoryObjects.OdataCount}";
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture) => throw new NotImplementedException();
}

public class DirectoryObjectsValueConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        var directoryObjects = (BaseCollectionPaginationCountResponse?)value;
        return directoryObjects switch
        {
            UserCollectionResponse => directoryObjects.GetValue<User>(),
            GroupCollectionResponse => directoryObjects.GetValue<Group>(),
            ApplicationCollectionResponse => directoryObjects.GetValue<Application>(),
            ServicePrincipalCollectionResponse => directoryObjects.GetValue<ServicePrincipal>(),
            DeviceCollectionResponse => directoryObjects.GetValue<Device>(),
            _ => Enumerable.Empty<DirectoryObject>()
        };
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
    {
        throw new NotImplementedException();
    }
}

public class CollectionToCommaSeparatedStringConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
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

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        => throw new NotImplementedException();
}
