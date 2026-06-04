namespace OutlookQuickMove
{
    /// <summary>
    /// Describes an Outlook data file (store) for display and filtering.
    /// Only the values needed by the UI are captured; the live <c>Store</c> COM object is
    /// released immediately after enumeration rather than retained on this instance.
    /// </summary>
    internal sealed class StoreCandidate
    {
        public StoreCandidate(string storeKey, string displayName, string displayText)
        {
            StoreKey = storeKey;
            DisplayName = displayName;
            DisplayText = displayText;
        }

        public string StoreKey { get; }

        public string DisplayName { get; }

        public string DisplayText { get; }

        public override string ToString()
        {
            return DisplayText;
        }
    }
}
