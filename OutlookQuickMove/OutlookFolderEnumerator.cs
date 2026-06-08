using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookQuickMove
{
    internal static class OutlookFolderEnumerator
    {
        private static readonly HashSet<string> ExcludedBranchNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "Search Folders",
            "RSS Feeds"
        };

        // A full folder enumeration opens (and releases) a large number of MAPI objects, so doing it
        // on every Quick Move / Go to Folder is wasteful and adds MAPI resource pressure. The result
        // holds no COM objects (only strings), so it is safe to cache and reuse for a short window.
        // The cache is invalidated when the enabled-store selection changes (signature mismatch),
        // when it ages past the TTL, or explicitly via InvalidateCache (e.g. after Settings save).
        private static readonly TimeSpan CacheLifetime = TimeSpan.FromMinutes(2);
        private static readonly TimeSpan DiskCacheLifetime = TimeSpan.FromMinutes(30);
        private static readonly object CacheGate = new object();
        private static FolderEnumerationResult cachedResult;
        private static string cachedSignature;
        private static DateTime cachedAtUtc;

        public static FolderEnumerationResult GetMailFolders(Outlook.Application application)
        {
            var overallStopwatch = Stopwatch.StartNew();
            var enabledStoreKeys = StoreFilterSettings.LoadEnabledStoreKeys();
            var signature = BuildSignature(enabledStoreKeys);

            lock (CacheGate)
            {
                if (cachedResult != null
                    && string.Equals(cachedSignature, signature, StringComparison.Ordinal)
                    && DateTime.UtcNow - cachedAtUtc < CacheLifetime)
                {
                    LogEnumerationCompleted("memory cache", cachedResult, overallStopwatch);
                    return cachedResult;
                }
            }

            var savedStoreEntries = StoreFilterSettings.LoadEnabledStoreEntries(enabledStoreKeys);
            var diskCacheStopwatch = Stopwatch.StartNew();
            long diskCacheAgeMs;
            var cachedFolders = FolderListCacheStore.Load(signature, DiskCacheLifetime, out diskCacheAgeMs);
            QuickMoveLog.Write("folder enumeration disk cache checked: folders=" + cachedFolders.Count
                + ", ageMs=" + diskCacheAgeMs
                + ", elapsedMs=" + diskCacheStopwatch.ElapsedMilliseconds + ".");

            FolderEnumerationResult result;
            string source;
            if (cachedFolders.Count > 0)
            {
                source = "disk cache";
                result = new FolderEnumerationResult(cachedFolders, new FolderEnumerationWarnings());
            }
            else if (savedStoreEntries.Count > 0)
            {
                source = "saved root identities";
                result = GetMailFoldersFromSavedStores(application, enabledStoreKeys, savedStoreEntries);
            }
            else
            {
                source = "legacy store scan";
                result = GetMailFolders(application, enabledStoreKeys);
            }

            if (cachedFolders.Count == 0 && result.Folders.Count > 0 && result.Warnings.Count == 0)
            {
                var saveStopwatch = Stopwatch.StartNew();
                FolderListCacheStore.Save(signature, result.Folders);
                QuickMoveLog.Write("folder enumeration disk cache save completed: folders=" + result.Folders.Count
                    + ", elapsedMs=" + saveStopwatch.ElapsedMilliseconds + ".");
            }
            else if (cachedFolders.Count == 0 && result.Folders.Count > 0)
            {
                QuickMoveLog.Write("folder enumeration disk cache save skipped: folders=" + result.Folders.Count
                    + ", warnings=" + result.Warnings.Count + ".");
            }

            lock (CacheGate)
            {
                cachedResult = result;
                cachedSignature = signature;
                cachedAtUtc = DateTime.UtcNow;
            }

            LogEnumerationCompleted(source, result, overallStopwatch);
            return result;
        }

        private static FolderEnumerationResult GetMailFoldersFromSavedStores(
            Outlook.Application application,
            HashSet<string> enabledStoreKeys,
            List<StoreFilterEntry> savedStoreEntries)
        {
            var totalStopwatch = Stopwatch.StartNew();
            var folders = new List<FolderCandidate>();
            var warnings = new FolderEnumerationWarnings();
            var fallbackStoreKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            if (application == null)
            {
                warnings.Add(FolderWarningKind.StoreUnreadable, "Outlook application is unavailable.");
                return new FolderEnumerationResult(folders, warnings);
            }

            Outlook.NameSpace session = null;
            try
            {
                QuickMoveLog.Write("saved-root folder enumeration session lookup started.");
                var sessionStopwatch = Stopwatch.StartNew();
                session = application.Session;
                QuickMoveLog.Write("saved-root folder enumeration session lookup completed: elapsedMs="
                    + sessionStopwatch.ElapsedMilliseconds + ".");
                QuickMoveLog.Write("saved-root folder enumeration started: selectedStores="
                    + savedStoreEntries.Count + ".");
                var seenKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (var entry in savedStoreEntries)
                {
                    var storeStopwatch = Stopwatch.StartNew();
                    var storeName = string.IsNullOrWhiteSpace(entry.DisplayName) ? entry.StoreKey : entry.DisplayName;
                    var countBeforeStore = folders.Count;
                    seenKeys.Add(entry.StoreKey);
                    Outlook.MAPIFolder root = null;
                    try
                    {
                        QuickMoveLog.Write("saved-root lookup started: store='" + storeName + "'.");
                        var rootLookupStopwatch = Stopwatch.StartNew();
                        root = session.GetFolderFromID(entry.RootEntryId, entry.RootStoreId);
                        QuickMoveLog.Write("saved-root lookup completed: store='" + storeName
                            + "', elapsedMs=" + rootLookupStopwatch.ElapsedMilliseconds + ".");
                        var displayName = string.IsNullOrWhiteSpace(entry.DisplayName)
                            ? GetFolderName(root)
                            : entry.DisplayName;
                        QuickMoveLog.Write("saved-root folder collection started: store='" + storeName + "'.");
                        var collectStopwatch = Stopwatch.StartNew();
                        CollectFolders(root, displayName, 0, false, folders, warnings);
                        QuickMoveLog.Write("saved-root folder collection completed: store='" + storeName
                            + "', addedFolders=" + (folders.Count - countBeforeStore)
                            + ", elapsedMs=" + collectStopwatch.ElapsedMilliseconds
                            + ", storeElapsedMs=" + storeStopwatch.ElapsedMilliseconds + ".");
                    }
                    catch (Exception ex)
                    {
                        fallbackStoreKeys.Add(entry.StoreKey);
                        QuickMoveLog.Write("saved root identity failed for selected store '" + storeName
                            + "'; legacy fallback will try this store.", ex);
                    }
                    finally
                    {
                        ComUtil.FinalRelease(root);
                    }
                }

                foreach (var enabledKey in enabledStoreKeys)
                {
                    if (!seenKeys.Contains(enabledKey))
                    {
                        fallbackStoreKeys.Add(enabledKey);
                        QuickMoveLog.Write("saved root identity is missing for selected store; legacy fallback will try store key: " + enabledKey);
                    }
                }
            }
            catch (Exception ex)
            {
                warnings.Add(FolderWarningKind.StoreUnreadable, "Could not resolve selected Outlook data files.");
                QuickMoveLog.Write("failed to resolve selected stores from saved root identities.", ex);
            }
            finally
            {
                ComUtil.FinalRelease(session);
                var reclaimStopwatch = Stopwatch.StartNew();
                ReclaimComResources();
                QuickMoveLog.Write("saved-root folder enumeration COM reclaim completed: elapsedMs="
                    + reclaimStopwatch.ElapsedMilliseconds + ".");
            }

            if (fallbackStoreKeys.Count > 0)
            {
                var fallbackStopwatch = Stopwatch.StartNew();
                QuickMoveLog.Write("saved-root legacy fallback started: selectedStores=" + fallbackStoreKeys.Count + ".");
                var fallbackResult = GetMailFolders(application, fallbackStoreKeys);
                folders.AddRange(fallbackResult.Folders);
                warnings.AddRange(fallbackResult.Warnings);

                QuickMoveLog.Write("saved-root legacy fallback completed: folders=" + fallbackResult.Folders.Count
                    + ", warnings=" + fallbackResult.Warnings.Count
                    + ", elapsedMs=" + fallbackStopwatch.ElapsedMilliseconds + ".");
            }

            folders.Sort((left, right) => string.Compare(left.DisplayPath, right.DisplayPath, StringComparison.OrdinalIgnoreCase));
            QuickMoveLog.Write("saved-root folder enumeration completed: folders=" + folders.Count
                + ", warnings=" + warnings.Count
                + ", elapsedMs=" + totalStopwatch.ElapsedMilliseconds + ".");
            return new FolderEnumerationResult(folders, warnings);
        }

        /// <summary>
        /// Drops the cached folder list so the next request re-enumerates. Call after anything that
        /// can change the available folders or the data-file selection (e.g. saving Settings).
        /// </summary>
        public static void InvalidateCache()
        {
            lock (CacheGate)
            {
                cachedResult = null;
                cachedSignature = null;
            }

            FolderListCacheStore.Clear();
        }

        public static FolderEnumerationResult GetMailFolders(Outlook.Application application, HashSet<string> enabledStoreKeys)
        {
            var totalStopwatch = Stopwatch.StartNew();
            var folders = new List<FolderCandidate>();
            var warnings = new FolderEnumerationWarnings();

            if (application == null)
            {
                warnings.Add(FolderWarningKind.StoreUnreadable, "Outlook application is unavailable.");
                return new FolderEnumerationResult(folders, warnings);
            }

            Outlook.NameSpace session = null;
            Outlook.Stores stores = null;
            try
            {
                QuickMoveLog.Write("legacy store-scan session lookup started.");
                var sessionStopwatch = Stopwatch.StartNew();
                session = application.Session;
                QuickMoveLog.Write("legacy store-scan session lookup completed: elapsedMs="
                    + sessionStopwatch.ElapsedMilliseconds + ".");
                QuickMoveLog.Write("legacy store-scan stores collection lookup started.");
                var storesStopwatch = Stopwatch.StartNew();
                stores = session.Stores;
                QuickMoveLog.Write("legacy store-scan stores collection lookup completed: elapsedMs="
                    + storesStopwatch.ElapsedMilliseconds + ".");
            }
            catch (Exception ex)
            {
                warnings.Add(FolderWarningKind.StoreUnreadable, "Could not read Outlook stores.");
                QuickMoveLog.Write("failed to read Outlook stores.", ex);
                ComUtil.FinalRelease(stores);
                ComUtil.FinalRelease(session);
                return new FolderEnumerationResult(folders, warnings);
            }

            try
            {
                QuickMoveLog.Write("legacy store-scan stores count read started.");
                var countStopwatch = Stopwatch.StartNew();
                int count = stores.Count;
                QuickMoveLog.Write("legacy store-scan stores count read completed: count=" + count
                    + ", elapsedMs=" + countStopwatch.ElapsedMilliseconds + ".");
                QuickMoveLog.Write("legacy store-scan folder enumeration started: filtered="
                    + (enabledStoreKeys != null && enabledStoreKeys.Count > 0) + ".");
                for (int i = 1; i <= count; i++)
                {
                    var storeStopwatch = Stopwatch.StartNew();
                    var countBeforeStore = folders.Count;
                    string displayName = null;
                    Outlook.Store store = null;
                    Outlook.MAPIFolder root = null;
                    try
                    {
                        QuickMoveLog.Write("legacy store-scan store item read started: index=" + i + ".");
                        var itemStopwatch = Stopwatch.StartNew();
                        store = stores[i];
                        QuickMoveLog.Write("legacy store-scan store item read completed: index=" + i
                            + ", elapsedMs=" + itemStopwatch.ElapsedMilliseconds + ".");
                        QuickMoveLog.Write("legacy store-scan file path read started: index=" + i + ".");
                        var filePathStopwatch = Stopwatch.StartNew();
                        var filePath = GetStoreFilePath(store);
                        QuickMoveLog.Write("legacy store-scan file path read completed: index=" + i
                            + ", hasFilePath=" + !string.IsNullOrWhiteSpace(filePath)
                            + ", elapsedMs=" + filePathStopwatch.ElapsedMilliseconds + ".");
                        if (enabledStoreKeys != null
                            && enabledStoreKeys.Count > 0
                            && !string.IsNullOrWhiteSpace(filePath)
                            && !StoreFilterSettings.IsStoreEnabled("file:" + filePath, enabledStoreKeys))
                        {
                            QuickMoveLog.Write("legacy store-scan skipped by file filter: index=" + i
                                + ", storeElapsedMs=" + storeStopwatch.ElapsedMilliseconds + ".");
                            continue;
                        }

                        QuickMoveLog.Write("legacy store-scan display name read started: index=" + i + ".");
                        var displayNameStopwatch = Stopwatch.StartNew();
                        displayName = GetStoreDisplayName(store, null);
                        QuickMoveLog.Write("legacy store-scan display name read completed: index=" + i
                            + ", store='" + displayName + "', elapsedMs="
                            + displayNameStopwatch.ElapsedMilliseconds + ".");
                        QuickMoveLog.Write("legacy store-scan store key read started: store='" + displayName + "'.");
                        var storeKeyStopwatch = Stopwatch.StartNew();
                        var storeKey = GetStoreKey(store, displayName, filePath);
                        QuickMoveLog.Write("legacy store-scan store key read completed: store='" + displayName
                            + "', elapsedMs=" + storeKeyStopwatch.ElapsedMilliseconds + ".");
                        if (!StoreFilterSettings.IsStoreEnabled(storeKey, enabledStoreKeys))
                        {
                            QuickMoveLog.Write("legacy store-scan skipped by store filter: store='" + displayName
                                + "', storeElapsedMs=" + storeStopwatch.ElapsedMilliseconds + ".");
                            continue;
                        }

                        QuickMoveLog.Write("legacy store-scan root folder read started: store='" + displayName + "'.");
                        var rootStopwatch = Stopwatch.StartNew();
                        root = store.GetRootFolder();
                        QuickMoveLog.Write("legacy store-scan root folder read completed: store='" + displayName
                            + "', elapsedMs=" + rootStopwatch.ElapsedMilliseconds + ".");
                        QuickMoveLog.Write("legacy store-scan folder collection started: store='" + displayName + "'.");
                        var collectStopwatch = Stopwatch.StartNew();
                        CollectFolders(root, displayName, 0, false, folders, warnings);
                        QuickMoveLog.Write("legacy store-scan folder collection completed: store='" + displayName
                            + "', addedFolders=" + (folders.Count - countBeforeStore)
                            + ", elapsedMs=" + collectStopwatch.ElapsedMilliseconds + ".");
                    }
                    catch (Exception ex)
                    {
                        var storeName = string.IsNullOrWhiteSpace(displayName)
                            ? GetStoreDisplayName(store, null)
                            : displayName;
                        warnings.Add(FolderWarningKind.StoreUnreadable, "Could not enumerate store: " + storeName);
                        QuickMoveLog.Write("failed to enumerate store '" + storeName + "'.", ex);
                    }
                    finally
                    {
                        ComUtil.FinalRelease(root);
                        ComUtil.FinalRelease(store);
                        QuickMoveLog.Write("legacy store-scan store finished: index=" + i
                            + ", addedFolders=" + (folders.Count - countBeforeStore)
                            + ", storeElapsedMs=" + storeStopwatch.ElapsedMilliseconds + ".");
                    }
                }
            }
            finally
            {
                ComUtil.FinalRelease(stores);
                ComUtil.FinalRelease(session);

                // The walk churned through many short-lived MAPI wrappers; reclaim promptly so the
                // resources are returned to MAPI instead of lingering until the next GC.
                var reclaimStopwatch = Stopwatch.StartNew();
                ReclaimComResources();
                QuickMoveLog.Write("legacy store-scan COM reclaim completed: elapsedMs="
                    + reclaimStopwatch.ElapsedMilliseconds + ".");
            }

            folders.Sort((left, right) => string.Compare(left.DisplayPath, right.DisplayPath, StringComparison.OrdinalIgnoreCase));
            QuickMoveLog.Write("legacy store-scan folder enumeration completed: folders=" + folders.Count
                + ", warnings=" + warnings.Count
                + ", elapsedMs=" + totalStopwatch.ElapsedMilliseconds + ".");
            return new FolderEnumerationResult(folders, warnings);
        }

        private static void LogEnumerationCompleted(string source, FolderEnumerationResult result, Stopwatch stopwatch)
        {
            if (result == null)
            {
                return;
            }

            QuickMoveLog.Write("folder enumeration completed from " + source
                + ": folders=" + result.Folders.Count
                + ", warnings=" + result.Warnings.Count
                + ", elapsedMs=" + stopwatch.ElapsedMilliseconds + ".");
        }

        private static string BuildSignature(HashSet<string> enabledStoreKeys)
        {
            if (enabledStoreKeys == null || enabledStoreKeys.Count == 0)
            {
                // Empty filter means "all stores"; distinguish it from an explicit selection.
                return "(all)";
            }

            return string.Join("\n", enabledStoreKeys.OrderBy(k => k, StringComparer.OrdinalIgnoreCase));
        }

        private static void ReclaimComResources()
        {
            try
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
            }
            catch (Exception ex)
            {
                QuickMoveLog.Write("failed to reclaim COM resources after enumeration.", ex);
            }
        }

        public static StoreEnumerationResult GetStores(Outlook.Application application)
        {
            var storesList = new List<StoreCandidate>();
            var errors = new List<string>();

            if (application == null)
            {
                errors.Add("Outlook application is unavailable.");
                return new StoreEnumerationResult(storesList, errors);
            }

            Outlook.NameSpace session = null;
            Outlook.Stores stores = null;
            try
            {
                session = application.Session;
                stores = session.Stores;
            }
            catch (Exception ex)
            {
                errors.Add("Could not read Outlook stores.");
                QuickMoveLog.Write("failed to read Outlook stores.", ex);
                ComUtil.FinalRelease(stores);
                ComUtil.FinalRelease(session);
                return new StoreEnumerationResult(storesList, errors);
            }

            try
            {
                int count = stores.Count;
                for (int i = 1; i <= count; i++)
                {
                    Outlook.Store store = null;
                    try
                    {
                        store = stores[i];
                        storesList.Add(CreateStoreCandidate(store));
                    }
                    catch (Exception ex)
                    {
                        var storeName = GetStoreDisplayName(store, null);
                        errors.Add("Could not read store: " + storeName);
                        QuickMoveLog.Write("failed to read store '" + storeName + "'.", ex);
                    }
                    finally
                    {
                        ComUtil.FinalRelease(store);
                    }
                }
            }
            finally
            {
                ComUtil.FinalRelease(stores);
                ComUtil.FinalRelease(session);
                ReclaimComResources();
            }

            storesList.Sort((left, right) => string.Compare(left.DisplayText, right.DisplayText, StringComparison.OrdinalIgnoreCase));
            return new StoreEnumerationResult(storesList, errors);
        }

        private static void CollectFolders(
            Outlook.MAPIFolder folder,
            string displayPath,
            int depth,
            bool excludedBranch,
            List<FolderCandidate> folders,
            FolderEnumerationWarnings warnings)
        {
            if (folder == null)
            {
                return;
            }

            var name = GetFolderName(folder);
            var isExcludedBranch = excludedBranch || ExcludedBranchNames.Contains(name);

            if (!isExcludedBranch && depth > 0 && IsMailDestination(folder))
            {
                var entryId = GetFolderEntryId(folder);
                var storeId = GetFolderStoreId(folder);
                if (!string.IsNullOrEmpty(entryId) && !string.IsNullOrEmpty(storeId))
                {
                    folders.Add(new FolderCandidate(displayPath, entryId, storeId));
                }
                else
                {
                    warnings.Add(FolderWarningKind.FolderIdentityMissing, "Could not identify folder: " + displayPath);
                }
            }

            Outlook.Folders childFolders;
            try
            {
                childFolders = folder.Folders;
            }
            catch (Exception ex)
            {
                warnings.Add(FolderWarningKind.SubfoldersUnreadable, "Could not read child folders under: " + displayPath);
                QuickMoveLog.Write("failed to read child folders under '" + displayPath + "'.", ex);
                return;
            }

            try
            {
                int count = childFolders.Count;
                for (int i = 1; i <= count; i++)
                {
                    Outlook.MAPIFolder child = null;
                    try
                    {
                        child = childFolders[i] as Outlook.MAPIFolder;
                        if (child == null)
                        {
                            continue;
                        }

                        var childName = GetFolderName(child);
                        var childPath = string.IsNullOrEmpty(displayPath) ? childName : displayPath + "\\" + childName;
                        CollectFolders(child, childPath, depth + 1, isExcludedBranch, folders, warnings);
                    }
                    catch (Exception ex)
                    {
                        warnings.Add(FolderWarningKind.SubfoldersUnreadable, "Could not enumerate a child folder under: " + displayPath);
                        QuickMoveLog.Write("failed to enumerate a child folder under '" + displayPath + "'.", ex);
                    }
                    finally
                    {
                        ComUtil.FinalRelease(child);
                    }
                }
            }
            finally
            {
                ComUtil.FinalRelease(childFolders);
            }
        }

        private static bool IsMailDestination(Outlook.MAPIFolder folder)
        {
            try
            {
                return folder.DefaultItemType == Outlook.OlItemType.olMailItem;
            }
            catch (Exception ex)
            {
                QuickMoveLog.Write("failed to read folder item type.", ex);
                return false;
            }
        }

        private static string GetStoreDisplayName(Outlook.Store store, Outlook.MAPIFolder root)
        {
            try
            {
                if (store != null && !string.IsNullOrWhiteSpace(store.DisplayName))
                {
                    return store.DisplayName;
                }
            }
            catch (Exception ex)
            {
                QuickMoveLog.Write("failed to read store display name.", ex);
            }

            return root == null ? "(unknown store)" : GetFolderName(root);
        }

        private static StoreCandidate CreateStoreCandidate(Outlook.Store store)
        {
            Outlook.MAPIFolder root = null;
            var displayName = GetStoreDisplayName(store, null);
            var filePath = GetStoreFilePath(store);
            var storeKey = GetStoreKey(store, displayName, filePath);
            var displayText = string.IsNullOrWhiteSpace(filePath)
                ? displayName
                : displayName + " (" + filePath + ")";

            try
            {
                root = store.GetRootFolder();
                return new StoreCandidate(
                    storeKey,
                    displayName,
                    displayText,
                    GetFolderEntryId(root),
                    GetFolderStoreId(root));
            }
            catch (Exception ex)
            {
                QuickMoveLog.Write("failed to read store root identity for '" + displayName + "'.", ex);
                return new StoreCandidate(storeKey, displayName, displayText, string.Empty, string.Empty);
            }
            finally
            {
                ComUtil.FinalRelease(root);
            }
        }

        private static string GetStoreKey(Outlook.Store store, string displayName, string filePath)
        {
            if (!string.IsNullOrWhiteSpace(filePath))
            {
                return "file:" + filePath;
            }

            try
            {
                if (store != null && !string.IsNullOrWhiteSpace(store.StoreID))
                {
                    return "store:" + store.StoreID;
                }
            }
            catch (Exception ex)
            {
                QuickMoveLog.Write("failed to read store id.", ex);
            }

            return "name:" + displayName;
        }

        private static string GetStoreFilePath(Outlook.Store store)
        {
            try
            {
                return store == null ? string.Empty : store.FilePath ?? string.Empty;
            }
            catch (Exception ex)
            {
                QuickMoveLog.Write("failed to read store file path.", ex);
                return string.Empty;
            }
        }

        private static string GetFolderName(Outlook.MAPIFolder folder)
        {
            try
            {
                return string.IsNullOrWhiteSpace(folder.Name) ? "(unnamed folder)" : folder.Name;
            }
            catch (Exception ex)
            {
                QuickMoveLog.Write("failed to read folder name.", ex);
                return "(unavailable folder)";
            }
        }

        private static string GetFolderEntryId(Outlook.MAPIFolder folder)
        {
            try
            {
                return folder.EntryID ?? string.Empty;
            }
            catch (Exception ex)
            {
                QuickMoveLog.Write("failed to read folder entry id.", ex);
                return string.Empty;
            }
        }

        private static string GetFolderStoreId(Outlook.MAPIFolder folder)
        {
            try
            {
                return folder.StoreID ?? string.Empty;
            }
            catch (Exception ex)
            {
                QuickMoveLog.Write("failed to read folder store id.", ex);
                return string.Empty;
            }
        }
    }
}
