using System;
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
                QuickMoveLog.Write("failed to release COM object.", ex);
            }
        }

        /// <summary>
        /// Releases a runtime callable wrapper completely (drives its reference count to zero in one
        /// call), for short-lived objects we exclusively own and want reclaimed promptly — notably
        /// the many folder/store wrappers opened during a full enumeration, which otherwise keep
        /// MAPI resources alive until garbage collection. Only use for RCWs that are not shared.
        /// </summary>
        public static void FinalRelease(object comObject)
        {
            if (comObject == null)
            {
                return;
            }

            try
            {
                if (Marshal.IsComObject(comObject))
                {
                    Marshal.FinalReleaseComObject(comObject);
                }
            }
            catch (Exception ex)
            {
                QuickMoveLog.Write("failed to final-release COM object.", ex);
            }
        }
    }
}
