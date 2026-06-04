using System.Collections.Generic;

namespace OutlookQuickMove
{
    internal sealed class FolderEnumerationResult
    {
        public FolderEnumerationResult(List<FolderCandidate> folders, List<string> errors)
        {
            Folders = folders;
            Errors = errors;
        }

        public List<FolderCandidate> Folders { get; }

        public List<string> Errors { get; }
    }
}
