// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using System;
using System.Linq;

namespace MsGraph_Samples
{
    public static class ExtensionMethods
    {
        public static bool In(this string x, params string[] items)
        {
            return items.Any(i => i.Trim().Equals(x, StringComparison.InvariantCultureIgnoreCase));
        }
    }
}
