using System;

namespace OutlookQuickMove
{
    /// <summary>
    /// One mail item moved by Quick Move, recorded so the move can be reversed later. The moved
    /// item is located by <see cref="MovedStoreId"/> + <see cref="MovedEntryId"/> (EntryIDs change
    /// on move, so this is the id in the destination); the original folder is resolved from
    /// <see cref="SourceFolderStoreId"/> + <see cref="SourceFolderEntryId"/>. The display paths and
    /// subject are kept for the history list, <see cref="OriginalUnRead"/> to restore the read
    /// state on rollback, <see cref="MovedAtUtc"/> for ordering/display/recency-based eviction, and
    /// <see cref="BatchId"/> to group the items moved together by one Quick Move action.
    /// </summary>
    internal sealed class UndoEntry
    {
        public UndoEntry(
            DateTime movedAtUtc,
            string movedEntryId,
            string movedStoreId,
            string sourceFolderEntryId,
            string sourceFolderStoreId,
            string sourceFolderPath,
            string targetFolderPath,
            bool originalUnRead,
            string subject,
            string batchId)
        {
            MovedAtUtc = movedAtUtc;
            MovedEntryId = movedEntryId ?? string.Empty;
            MovedStoreId = movedStoreId ?? string.Empty;
            SourceFolderEntryId = sourceFolderEntryId ?? string.Empty;
            SourceFolderStoreId = sourceFolderStoreId ?? string.Empty;
            SourceFolderPath = sourceFolderPath ?? string.Empty;
            TargetFolderPath = targetFolderPath ?? string.Empty;
            OriginalUnRead = originalUnRead;
            Subject = subject ?? string.Empty;
            BatchId = batchId ?? string.Empty;
        }

        public DateTime MovedAtUtc { get; }

        public string MovedEntryId { get; }

        public string MovedStoreId { get; }

        public string SourceFolderEntryId { get; }

        public string SourceFolderStoreId { get; }

        public string SourceFolderPath { get; }

        public string TargetFolderPath { get; }

        public bool OriginalUnRead { get; }

        public string Subject { get; }

        public string BatchId { get; }

        /// <summary>Identity of the moved item, used to select and de-duplicate history entries.</summary>
        public string Key
        {
            get { return BuildKey(MovedStoreId, MovedEntryId); }
        }

        public static string BuildKey(string movedStoreId, string movedEntryId)
        {
            return (movedStoreId ?? string.Empty) + "|" + (movedEntryId ?? string.Empty);
        }

        /// <summary>
        /// True only when enough identity is present to both find the moved item again and resolve
        /// its original folder. Entries without it are never recorded (the move still happens).
        /// </summary>
        public bool HasUndoIdentity
        {
            get
            {
                return !string.IsNullOrEmpty(MovedEntryId)
                    && !string.IsNullOrEmpty(MovedStoreId)
                    && !string.IsNullOrEmpty(SourceFolderEntryId)
                    && !string.IsNullOrEmpty(SourceFolderStoreId);
            }
        }

        public override string ToString()
        {
            var subject = string.IsNullOrWhiteSpace(Subject) ? "(no subject)" : Subject;
            var source = string.IsNullOrWhiteSpace(SourceFolderPath) ? "(unknown)" : SourceFolderPath;
            var target = string.IsNullOrWhiteSpace(TargetFolderPath) ? "(unknown)" : TargetFolderPath;
            return MovedAtUtc.ToLocalTime().ToString("yyyy-MM-dd HH:mm") + "  " + subject
                + "   [ " + source + " -> " + target + " ]";
        }
    }
}
