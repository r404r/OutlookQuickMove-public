using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace OutlookQuickMove
{
    internal static class StoreFilterSettings
    {
        private const string FilterFileName = "store-filter.txt";
        private const string RootFileName = "store-filter-roots.txt";
        private static readonly object RootGate = new object();

        public static HashSet<string> LoadEnabledStoreKeys()
        {
            var keys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var raw = LoadRawSettings();
            if (string.IsNullOrWhiteSpace(raw))
            {
                return keys;
            }

            foreach (var line in raw.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries))
            {
                var key = DecodeKey(line.Trim());
                if (!string.IsNullOrWhiteSpace(key))
                {
                    keys.Add(key);
                }
            }

            return keys;
        }

        /// <summary>
        /// The saved store root identities to enumerate. An empty filter means "all stores", so in
        /// that case every saved root is returned (this is what lets the default configuration use
        /// the saved-root fast path instead of binding the whole <c>Stores</c> collection). With an
        /// explicit subset, only the enabled keys are returned. Only entries with a usable root
        /// identity are included.
        /// </summary>
        public static List<StoreFilterEntry> LoadEnabledStoreEntries(HashSet<string> enabledStoreKeys)
        {
            var result = new List<StoreFilterEntry>();
            var filterIsAll = enabledStoreKeys == null || enabledStoreKeys.Count == 0;
            foreach (var entry in LoadAllStoreEntries())
            {
                if (!entry.HasRootIdentity)
                {
                    continue;
                }

                if (filterIsAll || enabledStoreKeys.Contains(entry.StoreKey))
                {
                    result.Add(entry);
                }
            }

            return result;
        }

        /// <summary>
        /// Whether the persisted root baseline contains at least one usable root. Used to decide
        /// whether a one-time bootstrap scan is needed — as opposed to the enable filter simply
        /// matching none of the saved roots, which must NOT trigger a full store scan.
        /// </summary>
        public static bool HasAnyStoreRoots()
        {
            foreach (var entry in LoadAllStoreEntries())
            {
                if (entry.HasRootIdentity)
                {
                    return true;
                }
            }

            return false;
        }

        private static List<StoreFilterEntry> LoadAllStoreEntries()
        {
            var entries = new List<StoreFilterEntry>();
            string raw;
            try
            {
                var path = QuickMovePaths.DataFile(RootFileName);
                if (!File.Exists(path))
                {
                    return entries;
                }

                raw = File.ReadAllText(path, Encoding.UTF8);
            }
            catch (Exception ex)
            {
                QuickMoveLog.Write("Failed to read store root filter settings file.", ex);
                return entries;
            }

            foreach (var line in raw.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries))
            {
                var entry = ParseEntry(line);
                if (entry != null)
                {
                    entries.Add(entry);
                }
            }

            return entries;
        }

        /// <summary>
        /// Persists the enabled store keys to the per-user settings file, which is the single
        /// source of truth (it survives VSTO deployment refreshes that reset per-version
        /// application settings). Returns false (and logs) when the file could not be written so
        /// callers can warn the user instead of silently losing the selection.
        /// </summary>
        public static bool SaveEnabledStoreKeys(IEnumerable<string> storeKeys)
        {
            var encodedKeys = storeKeys == null
                ? Enumerable.Empty<string>()
                : storeKeys
                    .Where(key => !string.IsNullOrWhiteSpace(key))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .Select(EncodeKey);

            var settingsText = string.Join(Environment.NewLine, encodedKeys);

            try
            {
                var settingsPath = GetSettingsPath();
                Directory.CreateDirectory(Path.GetDirectoryName(settingsPath));
                File.WriteAllText(settingsPath, settingsText, Encoding.UTF8);
                return true;
            }
            catch (Exception ex)
            {
                QuickMoveLog.Write("Failed to save store filter settings.", ex);
                return false;
            }
        }

        /// <summary>
        /// Replaces the saved root baseline with the given stores. Use only after a clean full scan;
        /// because it replaces the whole file, stores no longer present are pruned. Called with all
        /// known stores, not just the enabled subset, so the enable filter and the root baseline stay
        /// decoupled.
        /// </summary>
        public static bool SaveEnabledStoreEntries(IEnumerable<StoreCandidate> stores)
        {
            var entries = ToStoreEntries(stores);

            lock (RootGate)
            {
                return WriteStoreEntries(entries);
            }
        }

        /// <summary>
        /// Adds/updates roots from a partial store scan without pruning roots that were not observed.
        /// Use this when the scan reported errors or returned stores with missing root identities.
        /// </summary>
        public static bool MergeStoreRoots(IEnumerable<StoreCandidate> stores)
        {
            var entries = ToStoreEntries(stores).ToList();
            if (entries.Count == 0)
            {
                return true;
            }

            lock (RootGate)
            {
                var merged = LoadAllStoreEntries();
                foreach (var entry in entries)
                {
                    MergeEntry(merged, entry);
                }

                return WriteStoreEntries(merged);
            }
        }

        /// <summary>
        /// Adds or updates (by store key) a single store's root identity in the baseline, used by
        /// the <c>StoreAdd</c> tracker to pick up a newly mounted data file without a full scan.
        /// </summary>
        public static bool MergeStoreRoot(StoreFilterEntry entry)
        {
            if (entry == null || string.IsNullOrEmpty(entry.StoreKey) || !entry.HasRootIdentity)
            {
                return false;
            }

            lock (RootGate)
            {
                var merged = new List<StoreFilterEntry>();
                merged.AddRange(LoadAllStoreEntries());
                MergeEntry(merged, entry);
                return WriteStoreEntries(merged);
            }
        }

        private static IEnumerable<StoreFilterEntry> ToStoreEntries(IEnumerable<StoreCandidate> stores)
        {
            return (stores ?? Enumerable.Empty<StoreCandidate>())
                .Where(store => store != null
                    && !string.IsNullOrWhiteSpace(store.StoreKey)
                    && store.HasRootIdentity)
                .Select(store => new StoreFilterEntry(store.StoreKey, store.DisplayName, store.RootEntryId, store.RootStoreId));
        }

        private static void MergeEntry(List<StoreFilterEntry> entries, StoreFilterEntry entry)
        {
            if (entries == null || entry == null || string.IsNullOrWhiteSpace(entry.StoreKey) || !entry.HasRootIdentity)
            {
                return;
            }

            var replaced = false;
            for (var i = entries.Count - 1; i >= 0; i--)
            {
                if (string.Equals(entries[i].StoreKey, entry.StoreKey, StringComparison.OrdinalIgnoreCase))
                {
                    if (!replaced)
                    {
                        entries[i] = entry;
                        replaced = true;
                    }
                    else
                    {
                        entries.RemoveAt(i);
                    }
                }
            }

            if (!replaced)
            {
                entries.Add(entry);
            }
        }

        /// <summary>
        /// Removes a single store's root from the baseline (strict lazy prune when a saved root is
        /// definitively not found). Returns false when nothing matched.
        /// </summary>
        public static bool RemoveStoreRoot(string storeKey)
        {
            if (string.IsNullOrEmpty(storeKey))
            {
                return false;
            }

            lock (RootGate)
            {
                var kept = new List<StoreFilterEntry>();
                var removed = false;
                foreach (var existing in LoadAllStoreEntries())
                {
                    if (string.Equals(existing.StoreKey, storeKey, StringComparison.OrdinalIgnoreCase))
                    {
                        removed = true;
                    }
                    else
                    {
                        kept.Add(existing);
                    }
                }

                return removed && WriteStoreEntries(kept);
            }
        }

        private static bool WriteStoreEntries(IEnumerable<StoreFilterEntry> entries)
        {
            var builder = new StringBuilder();
            foreach (var entry in entries)
            {
                if (entry == null || string.IsNullOrWhiteSpace(entry.StoreKey) || !entry.HasRootIdentity)
                {
                    continue;
                }

                builder.Append(EncodeKey(entry.StoreKey)).Append('\t')
                    .Append(EncodeKey(entry.DisplayName)).Append('\t')
                    .Append(EncodeKey(entry.RootEntryId)).Append('\t')
                    .Append(EncodeKey(entry.RootStoreId)).Append('\n');
            }

            try
            {
                var settingsPath = QuickMovePaths.DataFile(RootFileName);
                Directory.CreateDirectory(Path.GetDirectoryName(settingsPath));
                File.WriteAllText(settingsPath, builder.ToString(), Encoding.UTF8);
                return true;
            }
            catch (Exception ex)
            {
                QuickMoveLog.Write("Failed to save store root filter settings.", ex);
                return false;
            }
        }

        public static bool IsStoreEnabled(string storeKey, HashSet<string> enabledStoreKeys)
        {
            if (string.IsNullOrEmpty(storeKey))
            {
                return false;
            }

            return enabledStoreKeys == null
                || enabledStoreKeys.Count == 0
                || enabledStoreKeys.Contains(storeKey);
        }

        // Store keys are Base64-encoded so values containing line breaks or other delimiter
        // characters survive the line-based file format. This is encoding for safe storage, not
        // encryption: the keys are local identifiers (file paths / store IDs), not secrets.
        private static string EncodeKey(string key)
        {
            return Convert.ToBase64String(Encoding.UTF8.GetBytes(key));
        }

        private static string LoadRawSettings()
        {
            var settingsPath = GetSettingsPath();
            try
            {
                if (File.Exists(settingsPath))
                {
                    return File.ReadAllText(settingsPath, Encoding.UTF8);
                }
            }
            catch (Exception ex)
            {
                QuickMoveLog.Write("Failed to read store filter settings file.", ex);
            }

            return string.Empty;
        }

        private static string GetSettingsPath()
        {
            return QuickMovePaths.DataFile(FilterFileName);
        }

        private static string DecodeKey(string encodedKey)
        {
            try
            {
                return Encoding.UTF8.GetString(Convert.FromBase64String(encodedKey));
            }
            catch (FormatException)
            {
                return encodedKey;
            }
        }

        private static StoreFilterEntry ParseEntry(string line)
        {
            try
            {
                var fields = line.Split('\t');
                if (fields.Length < 4)
                {
                    return null;
                }

                return new StoreFilterEntry(
                    DecodeKey(fields[0]),
                    DecodeKey(fields[1]),
                    DecodeKey(fields[2]),
                    DecodeKey(fields[3]));
            }
            catch (Exception ex)
            {
                QuickMoveLog.Write("Failed to parse a store root filter entry.", ex);
                return null;
            }
        }
    }
}
