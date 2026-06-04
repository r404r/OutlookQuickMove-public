using System;
using System.IO;

namespace OutlookQuickMove
{
    /// <summary>
    /// Resolves per-user file locations for Quick Move. The data folder lives under
    /// %APPDATA% so persisted preferences survive Outlook restarts and VSTO deployment
    /// refreshes (unlike per-version application settings, which are reset on refresh).
    /// </summary>
    internal static class QuickMovePaths
    {
        public static string DataFile(string fileName)
        {
            return Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "OutlookQuickMove",
                fileName);
        }
    }
}
