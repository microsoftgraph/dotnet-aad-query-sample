// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using System;
using System.IO;
using System.Reflection;
using System.Security.Cryptography;
using Microsoft.Identity.Client;

namespace MsGraph_Samples.Helpers
{
    static class TokenCacheHelper
    {
        private static readonly object FileLock = new object();
        private static readonly string LocalAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
        private static readonly string ProjectName = Assembly.GetCallingAssembly().GetName().Name ?? "tokencache";
        private static readonly string CacheFilePath = $"{LocalAppData}\\{ProjectName}\\msalcache.bin";

        public static void BeforeAccessNotification(TokenCacheNotificationArgs args)
        {
            if (!File.Exists(CacheFilePath))
                return;

            lock (FileLock)
            {
                var data = ProtectedData.Unprotect(File.ReadAllBytes(CacheFilePath), null, DataProtectionScope.CurrentUser);
                args.TokenCache.DeserializeMsalV3(data);
            }
        }

        public static void AfterAccessNotification(TokenCacheNotificationArgs args)
        {
            if (!args.HasStateChanged)
                return;

            // if the access operation resulted in a cache update
            lock (FileLock)
            {
                var cacheDirectory = Path.GetDirectoryName(CacheFilePath);
                if (!Directory.Exists(cacheDirectory))
                    Directory.CreateDirectory(cacheDirectory);

                var data = ProtectedData.Protect(args.TokenCache.SerializeMsalV3(), null, DataProtectionScope.CurrentUser);
                File.WriteAllBytes(CacheFilePath, data);
            }
        }

        internal static void EnableSerialization(ITokenCache tokenCache)
        {
            tokenCache.SetBeforeAccess(BeforeAccessNotification);
            tokenCache.SetAfterAccess(AfterAccessNotification);
        }
    }
}