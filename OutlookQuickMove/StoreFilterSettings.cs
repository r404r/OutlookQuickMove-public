using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace OutlookQuickMove
{
    internal static class StoreFilterSettings
    {
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
            return QuickMovePaths.DataFile("store-filter.txt");
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
    }
}
