using System;
using System.Diagnostics;
using System.IO;

namespace OutlookQuickMove
{
    /// <summary>
    /// Lightweight, best-effort diagnostic log that is available in Release builds, not just
    /// under DEBUG. Honors the project rule "never silently swallow failures": exceptions that
    /// are caught and handled are still recorded here for later diagnosis. Writes never throw,
    /// and the file is capped to avoid unbounded growth.
    /// </summary>
    internal static class QuickMoveLog
    {
        private const long MaxFileBytes = 1024 * 1024; // 1 MB
        private static readonly object Gate = new object();

        public static void Write(string message)
        {
            Write(message, null);
        }

        public static void Write(string message, Exception exception)
        {
            var line = DateTime.Now.ToString("O") + " " + message;
            if (exception != null)
            {
                line += " " + exception;
            }

            Debug.WriteLine("Quick Move: " + line);

            try
            {
                var path = GetLogPath();
                lock (Gate)
                {
                    TrimIfTooLarge(path);
                    File.AppendAllText(path, line + Environment.NewLine);
                }
            }
            catch
            {
                // Logging must never disrupt the add-in.
            }
        }

        private static void TrimIfTooLarge(string path)
        {
            try
            {
                var info = new FileInfo(path);
                if (info.Exists && info.Length > MaxFileBytes)
                {
                    File.Delete(path);
                }
            }
            catch
            {
            }
        }

        private static string GetLogPath()
        {
            return Path.Combine(Path.GetTempPath(), "OutlookQuickMove.log");
        }
    }
}
