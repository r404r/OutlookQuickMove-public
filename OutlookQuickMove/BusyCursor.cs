using System;
using System.Windows.Forms;

namespace OutlookQuickMove
{
    /// <summary>
    /// Shows a wait cursor for the duration of a synchronous, potentially slow operation
    /// (such as enumerating every folder in every store) so Outlook signals "busy" instead of
    /// appearing frozen. Deliberately minimal: the work still runs on the UI thread, so this is
    /// a feedback hint rather than background processing.
    /// </summary>
    internal sealed class BusyCursor : IDisposable
    {
        private readonly Cursor previousCursor;

        private BusyCursor()
        {
            previousCursor = Cursor.Current;
            Cursor.Current = Cursors.WaitCursor;
        }

        public static BusyCursor Show()
        {
            return new BusyCursor();
        }

        public void Dispose()
        {
            Cursor.Current = previousCursor;
        }
    }
}
