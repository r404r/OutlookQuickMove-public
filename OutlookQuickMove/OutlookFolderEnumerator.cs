using System;
using System.Collections.Generic;
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

        public static FolderEnumerationResult GetMailFolders(Outlook.Application application)
        {
            return GetMailFolders(application, StoreFilterSettings.LoadEnabledStoreKeys());
        }

        public static FolderEnumerationResult GetMailFolders(Outlook.Application application, HashSet<string> enabledStoreKeys)
        {
            var folders = new List<FolderCandidate>();
            var errors = new List<string>();

            if (application == null)
            {
                errors.Add("Outlook application is unavailable.");
                return new FolderEnumerationResult(folders, errors);
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
                ComUtil.Release(stores);
                ComUtil.Release(session);
                return new FolderEnumerationResult(folders, errors);
            }

            try
            {
                int count = stores.Count;
                for (int i = 1; i <= count; i++)
                {
                    Outlook.Store store = null;
                    Outlook.MAPIFolder root = null;
                    try
                    {
                        store = stores[i];
                        var displayName = GetStoreDisplayName(store, null);
                        var filePath = GetStoreFilePath(store);
                        var storeKey = GetStoreKey(store, displayName, filePath);

                        if (!StoreFilterSettings.IsStoreEnabled(storeKey, enabledStoreKeys))
                        {
                            continue;
                        }

                        root = store.GetRootFolder();
                        CollectFolders(root, displayName, 0, false, folders, errors);
                    }
                    catch (Exception ex)
                    {
                        var storeName = GetStoreDisplayName(store, null);
                        errors.Add("Could not enumerate store: " + storeName);
                        QuickMoveLog.Write("failed to enumerate store '" + storeName + "'.", ex);
                    }
                    finally
                    {
                        ComUtil.Release(root);
                        ComUtil.Release(store);
                    }
                }
            }
            finally
            {
                ComUtil.Release(stores);
                ComUtil.Release(session);
            }

            folders.Sort((left, right) => string.Compare(left.DisplayPath, right.DisplayPath, StringComparison.OrdinalIgnoreCase));
            return new FolderEnumerationResult(folders, errors);
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
                ComUtil.Release(stores);
                ComUtil.Release(session);
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
                        ComUtil.Release(store);
                    }
                }
            }
            finally
            {
                ComUtil.Release(stores);
                ComUtil.Release(session);
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
            List<string> errors)
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
                    errors.Add("Could not identify folder: " + displayPath);
                }
            }

            Outlook.Folders childFolders;
            try
            {
                childFolders = folder.Folders;
            }
            catch (Exception ex)
            {
                errors.Add("Could not read child folders under: " + displayPath);
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
                        CollectFolders(child, childPath, depth + 1, isExcludedBranch, folders, errors);
                    }
                    catch (Exception ex)
                    {
                        errors.Add("Could not enumerate a child folder under: " + displayPath);
                        QuickMoveLog.Write("failed to enumerate a child folder under '" + displayPath + "'.", ex);
                    }
                    finally
                    {
                        ComUtil.Release(child);
                    }
                }
            }
            finally
            {
                ComUtil.Release(childFolders);
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
            var displayName = GetStoreDisplayName(store, null);
            var filePath = GetStoreFilePath(store);
            var storeKey = GetStoreKey(store, displayName, filePath);
            var displayText = string.IsNullOrWhiteSpace(filePath)
                ? displayName
                : displayName + " (" + filePath + ")";

            return new StoreCandidate(storeKey, displayName, displayText);
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
