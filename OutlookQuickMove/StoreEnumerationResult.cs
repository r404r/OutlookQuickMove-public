using System.Collections.Generic;

namespace OutlookQuickMove
{
    internal sealed class StoreEnumerationResult
    {
        public StoreEnumerationResult(List<StoreCandidate> stores, List<string> errors)
        {
            Stores = stores;
            Errors = errors;
        }

        public List<StoreCandidate> Stores { get; }

        public List<string> Errors { get; }
    }
}
