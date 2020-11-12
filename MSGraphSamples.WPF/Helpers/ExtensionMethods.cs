// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using System;
using System.Linq;

namespace MsGraph_Samples
{
    public static class ExtensionMethods
    {
        public static bool In(this string s, params string[] items)
        {
            return items.Any(i => i.Trim().Equals(s, StringComparison.InvariantCultureIgnoreCase));
        }

        public static bool IsNullOrEmpty(this string? s) => string.IsNullOrEmpty(s);

        public static int NthIndexOf(this string input, char value, int nth, int startIndex = 0)
        {
            if (nth < 1)
                throw new ArgumentException("Input must be greater than 0", nameof(nth));
            if (nth == 1)
                return input.IndexOf(value, startIndex);

            return input.NthIndexOf(value, --nth, input.IndexOf(value, startIndex) + 1);
        }
    }
}