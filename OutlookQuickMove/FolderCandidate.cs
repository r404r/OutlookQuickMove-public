namespace OutlookQuickMove
{
    /// <summary>
    /// Describes a candidate destination folder by identity only.
    /// The folder is intentionally not held as a live COM object: the underlying
    /// <c>MAPIFolder</c> is resolved on demand via <c>Application.Session.GetFolderFromID</c>
    /// at move time. This avoids retaining hundreds of COM references for the lifetime of the
    /// dialog and tolerates the folder being renamed/moved while the dialog is open.
    /// </summary>
    internal sealed class FolderCandidate
    {
        public FolderCandidate(string displayPath, string entryId, string storeId)
        {
            DisplayPath = displayPath;
            EntryId = entryId;
            StoreId = storeId;
        }

        public string DisplayPath { get; }

        public string EntryId { get; }

        public string StoreId { get; }

        public override string ToString()
        {
            return DisplayPath;
        }
    }
}
