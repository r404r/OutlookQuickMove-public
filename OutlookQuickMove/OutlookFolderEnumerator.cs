using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
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

            var diskCacheStopwatch = Stopwatch.StartNew();
            long diskCacheAgeMs;
            var cachedFolders = FolderListCacheStore.Load(signature, DiskCacheLifetime, out diskCacheAgeMs);
            QuickMoveLog.WriteVerbose("folder enumeration disk cache checked: folders=" + cachedFolders.Count
                + ", ageMs=" + diskCacheAgeMs
                + ", elapsedMs=" + diskCacheStopwatch.ElapsedMilliseconds + ".");

            FolderEnumerationResult result;
            string source;
            if (cachedFolders.Count > 0)
            {
                source = "disk cache";
                result = new FolderEnumerationResult(cachedFolders, new FolderEnumerationWarnings());
            }
            else
            {
                var savedStoreEntries = StoreFilterSettings.LoadEnabledStoreEntries(enabledStoreKeys);
                StoreEnumerationResult bootstrapResult = null;

                // Bootstrap ONLY when no root baseline exists at all (fresh install, or an old build
                // that never wrote roots). A non-empty baseline that the current enable filter simply
                // doesn't match must NOT trigger a full session.Stores scan — that would reintroduce
                // the resource-heavy bind on the hot path. In that case return an empty list; the
                // user can rebuild via Refresh / Settings.
                if (savedStoreEntries.Count == 0 && !StoreFilterSettings.HasAnyStoreRoots())
                {
                    QuickMoveLog.Write("no saved store roots; bootstrapping the baseline from a one-time store scan.");
                    bootstrapResult = RefreshStoreRoots(application);
                    savedStoreEntries = StoreFilterSettings.LoadEnabledStoreEntries(enabledStoreKeys);
                }

                if (savedStoreEntries.Count > 0)
                {
                    source = "saved root identities";
                    result = GetMailFoldersFromSavedStores(application, savedStoreEntries);
                }
                else
                {
                    // Either no stores at all, or the enable filter matches none of the saved roots.
                    // Do not full-scan here; surface any bootstrap errors and let the user Refresh.
                    source = "no matching stores";
                    QuickMoveLog.Write("no folders enumerated: the enabled-store filter matched no saved roots (press Refresh to rebuild).");
                    var warnings = new FolderEnumerationWarnings();
                    AddStoreEnumerationErrors(warnings, bootstrapResult);
                    result = new FolderEnumerationResult(new List<FolderCandidate>(), warnings);
                }
            }

            if (cachedFolders.Count == 0 && result.Folders.Count > 0 && result.Warnings.Count == 0)
            {
                var saveStopwatch = Stopwatch.StartNew();
                FolderListCacheStore.Save(signature, result.Folders);
                QuickMoveLog.WriteVerbose("folder enumeration disk cache save completed: folders=" + result.Folders.Count
                    + ", elapsedMs=" + saveStopwatch.ElapsedMilliseconds + ".");
            }
            else if (cachedFolders.Count == 0 && result.Folders.Count > 0)
            {
                QuickMoveLog.WriteVerbose("folder enumeration disk cache save skipped: folders=" + result.Folders.Count
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

        /// <summary>
        /// Enumerates folders by opening each saved store root directly via
        /// <c>GetFolderFromID</c> — never binding the whole <c>Stores</c> collection. A root that
        /// fails to resolve is skipped with a single warning (no full-scan fallback); if the failure
        /// is a definitive not-found, that stale root is pruned from the baseline.
        /// </summary>
        private static FolderEnumerationResult GetMailFoldersFromSavedStores(
            Outlook.Application application,
            List<StoreFilterEntry> savedStoreEntries)
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
            try
            {
                session = application.Session;
                QuickMoveLog.WriteVerbose("saved-root folder enumeration started: stores=" + savedStoreEntries.Count + ".");
                foreach (var entry in savedStoreEntries)
                {
                    var storeStopwatch = Stopwatch.StartNew();
                    var storeName = string.IsNullOrWhiteSpace(entry.DisplayName) ? entry.StoreKey : entry.DisplayName;
                    var countBeforeStore = folders.Count;
                    Outlook.MAPIFolder root = null;
                    try
                    {
                        root = session.GetFolderFromID(entry.RootEntryId, entry.RootStoreId);
                        var displayName = string.IsNullOrWhiteSpace(entry.DisplayName)
                            ? GetFolderName(root)
                            : entry.DisplayName;
                        CollectFolders(root, displayName, 0, false, folders, warnings);
                        QuickMoveLog.WriteVerbose("saved-root store collected: store='" + storeName
                            + "', addedFolders=" + (folders.Count - countBeforeStore)
                            + ", storeElapsedMs=" + storeStopwatch.ElapsedMilliseconds + ".");
                    }
                    catch (Exception ex)
                    {
                        warnings.Add(FolderWarningKind.StoreUnreadable, "Could not open data file: " + storeName);
                        QuickMoveLog.Write("saved root identity failed for store '" + storeName + "'.", ex);

                        // Strict prune: only when the store is definitively gone, never on transient
                        // or resource errors (which would wipe the baseline during a resource storm).
                        if (IsNotFound(ex))
                        {
                            QuickMoveLog.Write("pruning stale store root (not found): store='" + storeName + "'.");
                            StoreFilterSettings.RemoveStoreRoot(entry.StoreKey);
                        }
                    }
                    finally
                    {
                        ComUtil.FinalRelease(root);
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
                ReclaimComResources();
            }

            folders.Sort((left, right) => string.Compare(left.DisplayPath, right.DisplayPath, StringComparison.OrdinalIgnoreCase));
            QuickMoveLog.WriteVerbose("saved-root folder enumeration completed: folders=" + folders.Count
                + ", warnings=" + warnings.Count
                + ", elapsedMs=" + totalStopwatch.ElapsedMilliseconds + ".");
            return new FolderEnumerationResult(folders, warnings);
        }

        /// <summary>
        /// Rebuilds the persisted store-root baseline from a full store scan (the user-initiated,
        /// retryable path used by Settings and Refresh). Clean scans replace the baseline and prune
        /// removed stores; partial scans merge successes and preserve existing roots for stores that
        /// could not be read. This is one of the few places that binds the whole <c>Stores</c>
        /// collection; the hot path does not.
        /// </summary>
        public static StoreEnumerationResult RefreshStoreRoots(Outlook.Application application)
        {
            try
            {
                var storeResult = GetStores(application);
                var hasScanErrors = storeResult.Errors.Count > 0;
                var saved = hasScanErrors
                    ? StoreFilterSettings.MergeStoreRoots(storeResult.Stores)
                    : StoreFilterSettings.SaveEnabledStoreEntries(storeResult.Stores);

                if (!saved)
                {
                    storeResult.Errors.Add("Could not save the Outlook data-file root baseline.");
                }

                QuickMoveLog.Write("store-root baseline refreshed: stores=" + storeResult.Stores.Count
                    + ", errors=" + storeResult.Errors.Count
                    + ", mode=" + (hasScanErrors ? "merge" : "replace")
                    + ".");
                return storeResult;
            }
            catch (Exception ex)
            {
                QuickMoveLog.Write("failed to rebuild store-root baseline.", ex);
                var errors = new List<string> { "Could not rebuild the Outlook data-file root baseline." };
                return new StoreEnumerationResult(new List<StoreCandidate>(), errors);
            }
        }

        private static void AddStoreEnumerationErrors(FolderEnumerationWarnings warnings, StoreEnumerationResult storeResult)
        {
            if (warnings == null || storeResult == null)
            {
                return;
            }

            foreach (var error in storeResult.Errors)
            {
                warnings.Add(FolderWarningKind.StoreUnreadable, error);
            }
        }

        /// <summary>
        /// Builds a <see cref="StoreFilterEntry"/> (key + display name + root identity) for one store,
        /// used by the <c>StoreAdd</c> tracker. Releases the root RCW; does not touch the caller's
        /// <c>Store</c>.
        /// </summary>
        internal static StoreFilterEntry CreateStoreFilterEntry(Outlook.Store store)
        {
            Outlook.MAPIFolder root = null;
            var displayName = GetStoreDisplayName(store, null);
            var filePath = GetStoreFilePath(store);
            var storeKey = GetStoreKey(store, displayName, filePath);
            try
            {
                root = store.GetRootFolder();
                return new StoreFilterEntry(storeKey, displayName, GetFolderEntryId(root), GetFolderStoreId(root));
            }
            catch (Exception ex)
            {
                QuickMoveLog.Write("failed to read store root identity for '" + displayName + "' (tracker).", ex);
                return new StoreFilterEntry(storeKey, displayName, string.Empty, string.Empty);
            }
            finally
            {
                ComUtil.FinalRelease(root);
            }
        }

        private static bool IsNotFound(Exception ex)
        {
            var com = ex as COMException;
            return com != null && unchecked((uint)com.ErrorCode) == 0x8004010Fu; // MAPI_E_NOT_FOUND
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
                        var candidate = CreateStoreCandidate(store);
                        storesList.Add(candidate);
                        if (!candidate.HasRootIdentity)
                        {
                            errors.Add("Could not read store root identity: " + candidate.DisplayName);
                        }
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
