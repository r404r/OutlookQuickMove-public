using System;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace OutlookQuickMove
{
    /// <summary>
    /// Helpers for releasing Outlook COM runtime callable wrappers (RCWs) deterministically.
    /// VSTO does not raise Shutdown reliably, so every COM object obtained during an
    /// operation must be released inline to avoid leaking references that keep Outlook alive.
    /// </summary>
    internal static class ComUtil
    {
        public static void Release(object comObject)
        {
            if (comObject == null)
            {
                return;
            }

            try
            {
                if (Marshal.IsComObject(comObject))
                {
                    Marshal.ReleaseComObject(comObject);
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine("Quick Move: failed to release COM object. " + ex);
            }
        }
    }
}
