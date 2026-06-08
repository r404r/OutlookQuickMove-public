namespace OutlookQuickMove
{
    /// <summary>
    /// Persisted identity for a selected Outlook data file. The store key remains the stable
    /// user-facing filter key; the root folder identity lets Quick Move enter the selected store
    /// directly without first touching every store in the Outlook profile.
    /// </summary>
    internal sealed class StoreFilterEntry
    {
        public StoreFilterEntry(string storeKey, string displayName, string rootEntryId, string rootStoreId)
        {
            StoreKey = storeKey ?? string.Empty;
            DisplayName = displayName ?? string.Empty;
            RootEntryId = rootEntryId ?? string.Empty;
            RootStoreId = rootStoreId ?? string.Empty;
        }

        public string StoreKey { get; }

        public string DisplayName { get; }

        public string RootEntryId { get; }

        public string RootStoreId { get; }

        public bool HasRootIdentity
        {
            get { return !string.IsNullOrEmpty(RootEntryId) && !string.IsNullOrEmpty(RootStoreId); }
        }
    }
}
