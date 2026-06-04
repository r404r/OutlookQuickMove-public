using System;
using System.IO;
using System.Text;

namespace OutlookQuickMove
{
    /// <summary>
    /// Persists the "Mark as read before moving" checkbox so its state survives closing the
    /// dialog and restarting Outlook. Stored next to the store filter in the per-user data
    /// folder. Best-effort: any failure is logged and the default (false) is used.
    /// </summary>
    internal static class MarkAsReadSetting
    {
        private const string FileName = "mark-as-read.txt";

        public static bool Load()
        {
            try
            {
                var path = QuickMovePaths.DataFile(FileName);
                if (File.Exists(path))
                {
                    var raw = File.ReadAllText(path, Encoding.UTF8).Trim();
                    return string.Equals(raw, "1", StringComparison.Ordinal)
                        || string.Equals(raw, "true", StringComparison.OrdinalIgnoreCase);
                }
            }
            catch (Exception ex)
            {
                QuickMoveLog.Write("Failed to read mark-as-read preference.", ex);
            }

            return false;
        }

        public static void Save(bool markAsRead)
        {
            try
            {
                var path = QuickMovePaths.DataFile(FileName);
                Directory.CreateDirectory(Path.GetDirectoryName(path));
                File.WriteAllText(path, markAsRead ? "1" : "0", Encoding.UTF8);
            }
            catch (Exception ex)
            {
                QuickMoveLog.Write("Failed to save mark-as-read preference.", ex);
            }
        }
    }
}
