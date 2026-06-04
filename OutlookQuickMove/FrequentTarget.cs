using System;

namespace OutlookQuickMove
{
    /// <summary>
    /// A move target folder the user has selected before, with how often and how recently it was
    /// used. Identified by <see cref="StoreId"/> + <see cref="EntryId"/> (the same identity used
    /// to resolve the folder at move time); <see cref="DisplayPath"/> is kept for display and is
    /// refreshed on each use so renames are reflected.
    /// </summary>
    internal sealed class FrequentTarget
    {
        public FrequentTarget(string storeId, string entryId, string displayPath, int count, DateTime lastUsedUtc)
        {
            StoreId = storeId ?? string.Empty;
            EntryId = entryId ?? string.Empty;
            DisplayPath = displayPath ?? string.Empty;
            Count = count;
            LastUsedUtc = lastUsedUtc;
        }

        public string StoreId { get; }

        public string EntryId { get; }

        public string DisplayPath { get; set; }

        public int Count { get; set; }

        public DateTime LastUsedUtc { get; set; }

        public string Key
        {
            get { return BuildKey(StoreId, EntryId); }
        }

        public static string BuildKey(string storeId, string entryId)
        {
            return (storeId ?? string.Empty) + "|" + (entryId ?? string.Empty);
        }

        public override string ToString()
        {
            var path = string.IsNullOrWhiteSpace(DisplayPath) ? "(unknown folder)" : DisplayPath;
            return path + "   (used " + Count + ")";
        }
    }
}
