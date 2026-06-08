using System.Collections.Generic;

namespace OutlookQuickMove
{
    /// <summary>
    /// The kinds of problems that can happen while enumerating folders, so the user-facing summary
    /// can report counts grouped by cause instead of one opaque total. Detailed per-item lines are
    /// still written to the log.
    /// </summary>
    internal enum FolderWarningKind
    {
        Other,
        StoreUnreadable,
        SubfoldersUnreadable,
        FolderIdentityMissing
    }

    /// <summary>
    /// Collects folder-enumeration warnings: the detailed messages (also logged) plus a count per
    /// <see cref="FolderWarningKind"/> for a categorized summary. Holds no COM objects.
    /// </summary>
    internal sealed class FolderEnumerationWarnings
    {
        private readonly List<string> messages = new List<string>();
        private readonly Dictionary<FolderWarningKind, int> counts = new Dictionary<FolderWarningKind, int>();

        public int Count
        {
            get { return messages.Count; }
        }

        public List<string> Messages
        {
            get { return messages; }
        }

        public IReadOnlyDictionary<FolderWarningKind, int> Counts
        {
            get { return counts; }
        }

        public void Add(FolderWarningKind kind, string message)
        {
            if (!string.IsNullOrEmpty(message))
            {
                messages.Add(message);
                QuickMoveLog.Write("folder enumeration warning (" + Describe(kind) + "): " + message);
            }

            int existing;
            counts.TryGetValue(kind, out existing);
            counts[kind] = existing + 1;
        }

        /// <summary>A short human label for a warning category, used in the summary dialog.</summary>
        public static string Describe(FolderWarningKind kind)
        {
            switch (kind)
            {
                case FolderWarningKind.StoreUnreadable:
                    return "Data files/stores that could not be read";
                case FolderWarningKind.SubfoldersUnreadable:
                    return "Folders whose subfolders could not be read";
                case FolderWarningKind.FolderIdentityMissing:
                    return "Folders with no usable identity";
                default:
                    return "Other folder warnings";
            }
        }
    }
}
