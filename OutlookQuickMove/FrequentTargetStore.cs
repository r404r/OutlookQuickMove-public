using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace OutlookQuickMove
{
    /// <summary>
    /// Persists the user's move-target usage history and the configurable cap on how many are
    /// surfaced. The history records counts for EVERY folder ever used (bounded only by a
    /// generous safety cap); the configurable number is a display limit, so a newly used folder
    /// accumulates real counts and can rise into the surfaced list instead of being starved out.
    /// Single source of truth is the per-user data folder, so the data survives Outlook restarts
    /// and VSTO deployment refreshes. All operations are best-effort: failures are logged and
    /// never thrown so they cannot disrupt a move.
    /// </summary>
    internal static class FrequentTargetStore
    {
        public const int DefaultMaxCount = 20;
        public const int MaxAllowedCount = 100;

        // Safety bound on how many records the history file keeps. Distinct destination folders
        // are naturally bounded by the folder tree; this is far above any display cap, so
        // eviction at this boundary is never visible in the surfaced list.
        private const int HistoryHardCap = 500;

        private const string ListFileName = "frequent-targets.txt";
        private const string MaxFileName = "frequent-targets-max.txt";

        public static int GetMaxCount()
        {
            try
            {
                var path = QuickMovePaths.DataFile(MaxFileName);
                if (File.Exists(path))
                {
                    var raw = File.ReadAllText(path, Encoding.UTF8).Trim();
                    if (int.TryParse(raw, out var value))
                    {
                        return Clamp(value);
                    }
                }
            }
            catch (Exception ex)
            {
                QuickMoveLog.Write("Failed to read frequent-target cap.", ex);
            }

            return DefaultMaxCount;
        }

        public static bool SaveMaxCount(int maxCount)
        {
            var clamped = Clamp(maxCount);
            try
            {
                var path = QuickMovePaths.DataFile(MaxFileName);
                Directory.CreateDirectory(Path.GetDirectoryName(path));
                File.WriteAllText(path, clamped.ToString(), Encoding.UTF8);
                return true;
            }
            catch (Exception ex)
            {
                QuickMoveLog.Write("Failed to save frequent-target cap.", ex);
                return false;
            }
        }

        /// <summary>
        /// The surfaced view: the most-used targets trimmed to the current display cap. Used by
        /// the Quick Move picker. Returns an empty list when the cap is 0 (feature off).
        /// </summary>
        public static List<FrequentTarget> LoadAll()
        {
            return SortAndTrim(LoadList(), GetMaxCount());
        }

        /// <summary>
        /// The full usage history (sorted most-used first, bounded only by the safety cap). Used
        /// by Settings for management, independent of the display cap.
        /// </summary>
        public static List<FrequentTarget> LoadHistory()
        {
            return SortAndTrim(LoadList(), HistoryHardCap);
        }

        /// <summary>
        /// Records one selection of a target folder into the history: increments its count (or
        /// adds it). The history retains counts for folders beyond the display cap so newcomers
        /// can accumulate enough uses to rise into the surfaced list. No-op when the feature is
        /// off (cap 0); existing history is preserved so re-enabling restores it.
        /// </summary>
        public static void RecordUse(string storeId, string entryId, string displayPath)
        {
            if (string.IsNullOrEmpty(storeId) || string.IsNullOrEmpty(entryId))
            {
                return;
            }

            if (GetMaxCount() <= 0)
            {
                // Feature disabled: stop recording, but keep any existing history.
                return;
            }

            var list = LoadList();
            var key = FrequentTarget.BuildKey(storeId, entryId);
            var existing = list.FirstOrDefault(t => string.Equals(t.Key, key, StringComparison.OrdinalIgnoreCase));
            if (existing != null)
            {
                existing.Count++;
                existing.LastUsedUtc = DateTime.UtcNow;
                if (!string.IsNullOrWhiteSpace(displayPath))
                {
                    existing.DisplayPath = displayPath;
                }
            }
            else
            {
                list.Add(new FrequentTarget(storeId, entryId, displayPath, 1, DateTime.UtcNow));
            }

            SaveList(SortAndTrim(list, HistoryHardCap));
        }

        /// <summary>
        /// Replaces the stored history with the given set (used by Settings after the user
        /// deletes or clears entries), bounded by the safety cap. Not trimmed to the display cap,
        /// so saving Settings never discards usage history beyond the surfaced window.
        /// </summary>
        public static bool ReplaceAll(IEnumerable<FrequentTarget> targets)
        {
            var list = targets == null ? new List<FrequentTarget>() : targets.ToList();
            return SaveList(SortAndTrim(list, HistoryHardCap));
        }

        private static int Clamp(int value)
        {
            if (value < 0)
            {
                return 0;
            }

            return value > MaxAllowedCount ? MaxAllowedCount : value;
        }

        private static List<FrequentTarget> SortAndTrim(List<FrequentTarget> list, int max)
        {
            var ordered = list
                .OrderByDescending(t => t.Count)
                .ThenByDescending(t => t.LastUsedUtc)
                .ThenBy(t => t.DisplayPath, StringComparer.OrdinalIgnoreCase)
                .ToList();

            if (max >= 0 && ordered.Count > max)
            {
                ordered = ordered.GetRange(0, max);
            }

            return ordered;
        }

        private static List<FrequentTarget> LoadList()
        {
            var list = new List<FrequentTarget>();
            string raw;
            try
            {
                var path = QuickMovePaths.DataFile(ListFileName);
                if (!File.Exists(path))
                {
                    return list;
                }

                raw = File.ReadAllText(path, Encoding.UTF8);
            }
            catch (Exception ex)
            {
                QuickMoveLog.Write("Failed to read frequent targets.", ex);
                return list;
            }

            foreach (var line in raw.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries))
            {
                var parsed = ParseLine(line);
                if (parsed != null)
                {
                    list.Add(parsed);
                }
            }

            return list;
        }

        private static FrequentTarget ParseLine(string line)
        {
            try
            {
                var fields = line.Split('\t');
                if (fields.Length < 5)
                {
                    return null;
                }

                if (!int.TryParse(fields[0], out var count) || !long.TryParse(fields[1], out var ticks))
                {
                    return null;
                }

                var storeId = Decode(fields[2]);
                var entryId = Decode(fields[3]);
                var displayPath = Decode(fields[4]);
                if (string.IsNullOrEmpty(storeId) || string.IsNullOrEmpty(entryId))
                {
                    return null;
                }

                return new FrequentTarget(storeId, entryId, displayPath, count, new DateTime(ticks, DateTimeKind.Utc));
            }
            catch (Exception ex)
            {
                QuickMoveLog.Write("Failed to parse a frequent-target record.", ex);
                return null;
            }
        }

        private static bool SaveList(List<FrequentTarget> list)
        {
            try
            {
                var builder = new StringBuilder();
                foreach (var target in list)
                {
                    builder.Append(target.Count).Append('\t')
                        .Append(target.LastUsedUtc.Ticks).Append('\t')
                        .Append(Encode(target.StoreId)).Append('\t')
                        .Append(Encode(target.EntryId)).Append('\t')
                        .Append(Encode(target.DisplayPath))
                        .Append('\n');
                }

                var path = QuickMovePaths.DataFile(ListFileName);
                Directory.CreateDirectory(Path.GetDirectoryName(path));
                File.WriteAllText(path, builder.ToString(), Encoding.UTF8);
                return true;
            }
            catch (Exception ex)
            {
                QuickMoveLog.Write("Failed to save frequent targets.", ex);
                return false;
            }
        }

        private static string Encode(string value)
        {
            return Convert.ToBase64String(Encoding.UTF8.GetBytes(value ?? string.Empty));
        }

        private static string Decode(string value)
        {
            try
            {
                return Encoding.UTF8.GetString(Convert.FromBase64String(value));
            }
            catch (FormatException)
            {
                return string.Empty;
            }
        }
    }
}
