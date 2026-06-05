using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace OutlookQuickMove
{
    /// <summary>
    /// Persists the recent Quick Move history so moves can be undone, plus the configurable cap on
    /// how many moved items are remembered. The cap bounds the stored history by recency (newest
    /// kept, oldest evicted) and is applied only when recording; reads return whatever is stored so
    /// turning the feature off (cap 0) stops new recording but keeps existing history undoable until
    /// it is cleared. Single source of truth is the per-user data folder, so the history survives
    /// Outlook restarts and VSTO deployment refreshes. All operations are best-effort: failures are
    /// logged and never thrown so they cannot disrupt a move or an undo.
    /// </summary>
    internal static class UndoStore
    {
        public const int DefaultMaxCount = 50;
        public const int MaxAllowedCount = 500;

        private const string ListFileName = "undo-history.txt";
        private const string MaxFileName = "undo-max.txt";

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
                QuickMoveLog.Write("Failed to read undo-history cap.", ex);
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
                QuickMoveLog.Write("Failed to save undo-history cap.", ex);
                return false;
            }
        }

        /// <summary>The full stored history, newest first. Drives the Undo dialog.</summary>
        public static List<UndoEntry> LoadAll()
        {
            return SortAndTrim(LoadList(), MaxAllowedCount);
        }

        /// <summary>Whether any undo history exists. Drives the ribbon button's enabled state.</summary>
        public static bool HasAny()
        {
            return LoadList().Count > 0;
        }

        /// <summary>
        /// Records the items moved by one Quick Move action, trimming the stored history to the
        /// current cap (oldest evicted). No-op when the feature is off (cap 0) or nothing has the
        /// identity needed to undo; existing history is preserved either way.
        /// </summary>
        public static void Append(IEnumerable<UndoEntry> entries)
        {
            if (entries == null)
            {
                return;
            }

            var toAdd = entries.Where(e => e != null && e.HasUndoIdentity).ToList();
            if (toAdd.Count == 0)
            {
                return;
            }

            if (GetMaxCount() <= 0)
            {
                // Feature disabled: stop recording, but keep any existing history.
                return;
            }

            var list = LoadList();
            list.AddRange(toAdd);
            SaveList(SortAndTrim(list, GetMaxCount()));
        }

        /// <summary>
        /// Replaces the stored history with the given set (used after a rollback removes the
        /// undone entries, or when the user clears history), bounded only by the safety cap so a
        /// smaller display cap never silently discards still-undoable entries.
        /// </summary>
        public static bool Save(IEnumerable<UndoEntry> entries)
        {
            var list = entries == null ? new List<UndoEntry>() : entries.ToList();
            return SaveList(SortAndTrim(list, MaxAllowedCount));
        }

        private static int Clamp(int value)
        {
            if (value < 0)
            {
                return 0;
            }

            return value > MaxAllowedCount ? MaxAllowedCount : value;
        }

        private static List<UndoEntry> SortAndTrim(List<UndoEntry> list, int max)
        {
            var ordered = list
                .OrderByDescending(e => e.MovedAtUtc)
                .ThenBy(e => e.Subject, StringComparer.OrdinalIgnoreCase)
                .ToList();

            if (max >= 0 && ordered.Count > max)
            {
                ordered = ordered.GetRange(0, max);
            }

            return ordered;
        }

        private static List<UndoEntry> LoadList()
        {
            var list = new List<UndoEntry>();
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
                QuickMoveLog.Write("Failed to read undo history.", ex);
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

        private static UndoEntry ParseLine(string line)
        {
            try
            {
                var fields = line.Split('\t');
                if (fields.Length < 9)
                {
                    return null;
                }

                if (!long.TryParse(fields[0], out var ticks) || !int.TryParse(fields[1], out var unread))
                {
                    return null;
                }

                var movedEntryId = Decode(fields[2]);
                var movedStoreId = Decode(fields[3]);
                var sourceEntryId = Decode(fields[4]);
                var sourceStoreId = Decode(fields[5]);
                var sourcePath = Decode(fields[6]);
                var targetPath = Decode(fields[7]);
                var subject = Decode(fields[8]);
                var batchId = fields.Length >= 10 ? Decode(fields[9]) : string.Empty;

                if (string.IsNullOrEmpty(movedEntryId) || string.IsNullOrEmpty(movedStoreId))
                {
                    return null;
                }

                return new UndoEntry(
                    new DateTime(ticks, DateTimeKind.Utc),
                    movedEntryId,
                    movedStoreId,
                    sourceEntryId,
                    sourceStoreId,
                    sourcePath,
                    targetPath,
                    unread != 0,
                    subject,
                    batchId);
            }
            catch (Exception ex)
            {
                QuickMoveLog.Write("Failed to parse an undo-history record.", ex);
                return null;
            }
        }

        private static bool SaveList(List<UndoEntry> list)
        {
            try
            {
                var builder = new StringBuilder();
                foreach (var entry in list)
                {
                    builder.Append(entry.MovedAtUtc.Ticks).Append('\t')
                        .Append(entry.OriginalUnRead ? 1 : 0).Append('\t')
                        .Append(Encode(entry.MovedEntryId)).Append('\t')
                        .Append(Encode(entry.MovedStoreId)).Append('\t')
                        .Append(Encode(entry.SourceFolderEntryId)).Append('\t')
                        .Append(Encode(entry.SourceFolderStoreId)).Append('\t')
                        .Append(Encode(entry.SourceFolderPath)).Append('\t')
                        .Append(Encode(entry.TargetFolderPath)).Append('\t')
                        .Append(Encode(entry.Subject)).Append('\t')
                        .Append(Encode(entry.BatchId))
                        .Append('\n');
                }

                var path = QuickMovePaths.DataFile(ListFileName);
                Directory.CreateDirectory(Path.GetDirectoryName(path));
                File.WriteAllText(path, builder.ToString(), Encoding.UTF8);
                return true;
            }
            catch (Exception ex)
            {
                QuickMoveLog.Write("Failed to save undo history.", ex);
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
