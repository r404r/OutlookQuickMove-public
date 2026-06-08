using System.Collections.Generic;

namespace OutlookQuickMove
{
    internal sealed class FolderEnumerationResult
    {
        public FolderEnumerationResult(List<FolderCandidate> folders, FolderEnumerationWarnings warnings)
        {
            Folders = folders;
            Warnings = warnings ?? new FolderEnumerationWarnings();
        }

        public List<FolderCandidate> Folders { get; }

        public FolderEnumerationWarnings Warnings { get; }
    }
}
