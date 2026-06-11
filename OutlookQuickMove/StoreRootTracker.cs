using System;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookQuickMove
{
    /// <summary>
    /// Keeps the persisted store-root baseline fresh for newly mounted data files, without ever
    /// scanning the whole <c>Stores</c> collection on the hot path. Subscribes once to
    /// <c>Stores.StoreAdd</c>; when a store mounts, it records that one store's root identity.
    ///
    /// This is the single deliberate exception to the "release every COM object" rule: the
    /// <c>Stores</c> collection RCW is retained for the add-in lifetime so the event stays live. It
    /// is one RCW and one MAPI notification — negligible, and far cheaper than the per–Quick Move
    /// full store bind it removes. The tracker is used ONLY to add roots, never to infer removals
    /// (absence of <c>StoreAdd</c> does not mean a store was removed); removal is reconciled by
    /// Settings/Refresh and the strict not-found prune.
    /// </summary>
    internal static class StoreRootTracker
    {
        private static readonly object Gate = new object();
        private static Outlook.Stores trackedStores;
        private static Outlook.StoresEvents_12_StoreAddEventHandler handler;

        public static void Start(Outlook.Application application)
        {
            if (application == null)
            {
                return;
            }

            try
            {
                lock (Gate)
                {
                    if (trackedStores != null)
                    {
                        return;
                    }

                    Outlook.NameSpace session = null;
                    try
                    {
                        session = application.Session;
                        trackedStores = session.Stores;
                    }
                    finally
                    {
                        ComUtil.Release(session);
                    }

                    handler = OnStoreAdd;
                    trackedStores.StoreAdd += handler;
                    QuickMoveLog.Write("store-add tracker started.");
                }
            }
            catch (Exception ex)
            {
                QuickMoveLog.Write("failed to start store-added tracker.", ex);
            }
        }

        public static void Stop()
        {
            try
            {
                lock (Gate)
                {
                    if (trackedStores != null && handler != null)
                    {
                        trackedStores.StoreAdd -= handler;
                    }

                    handler = null;
                    ComUtil.Release(trackedStores);
                    trackedStores = null;
                }
            }
            catch (Exception ex)
            {
                QuickMoveLog.Write("failed to stop store-added tracker.", ex);
            }
        }

        private static void OnStoreAdd(Outlook.Store store)
        {
            // The event-provided Store RCW enters our code, so we release it in finally per the COM
            // rule (the only retained reference is the single Stores collection). Roots opened inside
            // CreateStoreFilterEntry are released there.
            try
            {
                if (store == null)
                {
                    return;
                }

                var entry = OutlookFolderEnumerator.CreateStoreFilterEntry(store);
                if (entry == null || !entry.HasRootIdentity)
                {
                    QuickMoveLog.Write("store added but its root identity was unavailable; not merged.");
                    return;
                }

                if (StoreFilterSettings.MergeStoreRoot(entry))
                {
                    // Only drop the caches once the new root is actually persisted, so the next
                    // enumeration genuinely picks it up.
                    OutlookFolderEnumerator.InvalidateCache();
                    QuickMoveLog.Write("store added; merged root for '" + entry.DisplayName + "'.");
                }
                else
                {
                    QuickMoveLog.Write("store added but persisting its root failed; not merged: '" + entry.DisplayName + "'.");
                }
            }
            catch (Exception ex)
            {
                QuickMoveLog.Write("failed to handle StoreAdd.", ex);
            }
            finally
            {
                ComUtil.Release(store);
            }
        }
    }
}
