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
        private static readonly string CacheDirectoryPath = $"{LocalAppData}\\{ProjectName}\\";
        private static readonly string CacheFileName = "msalcache.bin";
        private static string CacheFilePath => Path.Combine(CacheDirectoryPath, CacheFileName);

        private static void BeforeAccessNotification(TokenCacheNotificationArgs args)
        {
            lock (FileLock)
            {
                args.TokenCache.DeserializeMsalV3(File.Exists(CacheFilePath)
                        ? ProtectedData.Unprotect(File.ReadAllBytes(CacheFilePath),
                                                  null,
                                                  DataProtectionScope.CurrentUser)
                        : null);
            }
        }

        private static void AfterAccessNotification(TokenCacheNotificationArgs args)
        {
            // if the access operation resulted in a cache update
            if (args.HasStateChanged)
            {
                lock (FileLock)
                {
                    // reflect changes in the persistent store
                    File.WriteAllBytes(CacheFilePath,
                                        ProtectedData.Protect(args.TokenCache.SerializeMsalV3(),
                                                                null,
                                                                DataProtectionScope.CurrentUser)
                                        );
                }
            }
        }
        public static void Clear()
        {
            lock (FileLock)
            {
                File.Delete(CacheFilePath);
            }
        }

        internal static void EnableSerialization(ITokenCache tokenCache)
        {
            tokenCache.SetBeforeAccess(BeforeAccessNotification);
            tokenCache.SetAfterAccess(AfterAccessNotification);
        }
    }
}