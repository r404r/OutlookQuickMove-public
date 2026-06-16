using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookQuickMove
{
    [ComVisible(true)]
    public sealed class QuickMoveRibbon : Office.IRibbonExtensibility
    {
        private const string UndoButtonId = "QuickMoveUndoButton";
        private const string QuickMoveImageId = "QuickMove32";
        private const string GoToFolderImageId = "GoToFolder32";
        private const string GoToMailFolderImageId = "GoToMailFolder32";
        private static readonly Dictionary<string, string> RibbonImageResources = new Dictionary<string, string>(StringComparer.Ordinal)
        {
            { "QuickMove16", "OutlookQuickMove.Assets.QuickMove16.png" },
            { QuickMoveImageId, "OutlookQuickMove.Assets.QuickMove32.png" },
            { "GoToFolder16", "OutlookQuickMove.Assets.GoToFolder16.png" },
            { GoToFolderImageId, "OutlookQuickMove.Assets.GoToFolder32.png" },
            { "GoToMailFolder16", "OutlookQuickMove.Assets.GoToMailFolder16.png" },
            { GoToMailFolderImageId, "OutlookQuickMove.Assets.GoToMailFolder32.png" },
        };

        // The ribbon UI, captured on load so the Undo button's enabled state can be refreshed
        // after a move or an undo. Static because moves run through static helpers.
        private static Office.IRibbonUI ribbonUi;

        public string GetCustomUI(string ribbonId)
        {
            ThisAddIn.DebugLog("GetCustomUI ribbonId=" + (ribbonId ?? "(null)"));
            if (!string.Equals(ribbonId, "Microsoft.Outlook.Explorer", StringComparison.OrdinalIgnoreCase))
            {
                return string.Empty;
            }

            return @"<?xml version=""1.0"" encoding=""UTF-8""?>
<customUI xmlns=""http://schemas.microsoft.com/office/2009/07/customui"" onLoad=""OnRibbonLoad"" loadImage=""LoadImage"">
  <ribbon>
    <tabs>
      <tab id=""OutlookQuickMoveTab"" label=""Quick Move"" insertAfterMso=""TabMail"">
        <group id=""OutlookQuickMoveActionsGroup"" label=""Actions"">
          <button id=""QuickMoveButton""
                  label=""Quick Move""
                  size=""large""
                  image=""" + QuickMoveImageId + @"""
                  onAction=""OnQuickMove"" />
          <button id=""QuickJumpButton""
                  label=""Go to Folder""
                  size=""large""
                  image=""" + GoToFolderImageId + @"""
                  onAction=""OnGoToFolder"" />
          <button id=""QuickJumpSelectedMailFolderButton""
                  label=""Go to Mail Folder""
                  size=""large""
                  image=""" + GoToMailFolderImageId + @"""
                  onAction=""OnGoToSelectedMailFolder"" />
          <button id=""QuickMoveUndoButton""
                  label=""Undo Quick Move...""
                  size=""large""
                  imageMso=""Undo""
                  getEnabled=""OnGetUndoEnabled""
                  onAction=""OnUndo"" />
          <button id=""QuickMoveSettingsButton""
                  label=""Settings""
                  size=""large""
                  imageMso=""ApplicationOptionsDialog""
                  onAction=""OnSettings"" />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";
        }

        public stdole.IPictureDisp LoadImage(string imageId)
        {
            try
            {
                string resourceName;
                if (string.IsNullOrEmpty(imageId) || !RibbonImageResources.TryGetValue(imageId, out resourceName))
                {
                    QuickMoveLog.Write("Ribbon image id was not recognized: " + (imageId ?? "(null)"));
                    return null;
                }

                using (var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(resourceName))
                {
                    if (stream == null)
                    {
                        QuickMoveLog.Write("Ribbon image resource was not found: " + resourceName);
                        return null;
                    }

                    using (var image = Image.FromStream(stream))
                    {
                        return PictureDispConverter.ToPictureDisp(image);
                    }
                }
            }
            catch (Exception ex)
            {
                QuickMoveLog.Write("Failed to load ribbon image: " + (imageId ?? "(null)"), ex);
                return null;
            }
        }

        public void OnQuickMove(Office.IRibbonControl control)
        {
            try
            {
                ExecuteQuickMove();
            }
            catch (Exception ex)
            {
                QuickMoveLog.Write("unexpected error.", ex);
                MessageBox.Show("Quick Move failed unexpectedly. Check the Quick Move log for details.", "Quick Move", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void OnGoToFolder(Office.IRibbonControl control)
        {
            try
            {
                ExecuteGoToFolder();
            }
            catch (Exception ex)
            {
                QuickMoveLog.Write("go to folder failed unexpectedly.", ex);
                MessageBox.Show("Go to Folder failed unexpectedly. Check the Quick Move log for details.", "Go to Folder", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void OnGoToSelectedMailFolder(Office.IRibbonControl control)
        {
            try
            {
                ExecuteGoToSelectedMailFolder();
            }
            catch (Exception ex)
            {
                QuickMoveLog.Write("go to mail folder failed unexpectedly.", ex);
                MessageBox.Show("Go to Mail Folder failed unexpectedly. Check the Quick Move log for details.", "Go to Mail Folder", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void OnSettings(Office.IRibbonControl control)
        {
            try
            {
                ExecuteSettings();
            }
            catch (Exception ex)
            {
                QuickMoveLog.Write("settings failed unexpectedly.", ex);
                MessageBox.Show("Quick Move settings failed unexpectedly. Check the Quick Move log for details.", "Quick Move", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void OnRibbonLoad(Office.IRibbonUI ribbonUI)
        {
            ribbonUi = ribbonUI;
        }

        public bool OnGetUndoEnabled(Office.IRibbonControl control)
        {
            try
            {
                return UndoStore.HasAny();
            }
            catch (Exception ex)
            {
                QuickMoveLog.Write("failed to evaluate undo availability.", ex);
                return false;
            }
        }

        public void OnUndo(Office.IRibbonControl control)
        {
            try
            {
                ExecuteUndo();
            }
            catch (Exception ex)
            {
                QuickMoveLog.Write("undo failed unexpectedly.", ex);
                MessageBox.Show("Undo failed unexpectedly. Check the Quick Move log for details.", "Quick Move", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private static void InvalidateUndoButton()
        {
            try
            {
                if (ribbonUi != null)
                {
                    ribbonUi.InvalidateControl(UndoButtonId);
                }
            }
            catch (Exception ex)
            {
                QuickMoveLog.Write("failed to refresh the undo button.", ex);
            }
        }

        private sealed class PictureDispConverter : AxHost
        {
            private PictureDispConverter()
                : base(null)
            {
            }

            public static stdole.IPictureDisp ToPictureDisp(Image image)
            {
                return (stdole.IPictureDisp)GetIPictureDispFromPicture(image);
            }
        }

        private static void ExecuteQuickMove()
        {
            var application = Globals.ThisAddIn.Application;
            Outlook.Explorer explorer = null;
            Outlook.Selection selection = null;
            var selectedMail = new List<Outlook.MailItem>();
            var skippedNonMail = 0;

            try
            {
                explorer = application.ActiveExplorer();
                if (explorer == null)
                {
                    MessageBox.Show("No active Outlook explorer window is available.", "Quick Move", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                selection = explorer.Selection;
                if (selection == null || selection.Count == 0)
                {
                    MessageBox.Show("Select one or more mail items before using Quick Move.", "Quick Move", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                for (int i = selection.Count; i >= 1; i--)
                {
                    object selectedItem = null;
                    try
                    {
                        selectedItem = selection[i];
                    }
                    catch (Exception ex)
                    {
                        skippedNonMail++;
                        QuickMoveLog.Write("failed to read selected item at index " + i + ".", ex);
                        continue;
                    }

                    var mail = selectedItem as Outlook.MailItem;
                    if (mail == null)
                    {
                        skippedNonMail++;
                        ComUtil.Release(selectedItem);
                        continue;
                    }

                    selectedMail.Add(mail);
                }

                if (selectedMail.Count == 0)
                {
                    MessageBox.Show("The current selection does not contain mail items.", "Quick Move", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                FolderEnumerationResult folderResult;
                using (BusyCursor.Show())
                {
                    folderResult = OutlookFolderEnumerator.GetMailFolders(application);
                }

                if (folderResult.Folders.Count == 0)
                {
                    MessageBox.Show("No suitable mail folders were found.", "Quick Move", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                string targetEntryId;
                string targetStoreId;
                bool markAsRead;
                FolderEnumerationWarnings folderWarnings;
                using (var form = new FolderPickerForm(
                    folderResult.Folders,
                    FolderPickerOptions.ForQuickMove(),
                    folderResult.Warnings,
                    delegate { return RefreshFolderList(application); }))
                {
                    if (form.ShowDialog() != DialogResult.OK)
                    {
                        return;
                    }

                    targetEntryId = form.SelectedFolderEntryId;
                    targetStoreId = form.SelectedFolderStoreId;
                    markAsRead = form.MarkAsReadBeforeMoving;
                    folderWarnings = form.FolderWarnings;
                }

                if (string.IsNullOrEmpty(targetEntryId) || string.IsNullOrEmpty(targetStoreId))
                {
                    MessageBox.Show("Select a target folder before moving mail.", "Quick Move", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                MoveSelectedMail(application, selectedMail, targetEntryId, targetStoreId, markAsRead, skippedNonMail, folderWarnings);
            }
            finally
            {
                foreach (var mail in selectedMail)
                {
                    ComUtil.Release(mail);
                }

                ComUtil.Release(selection);
                ComUtil.Release(explorer);
            }
        }

        private static void MoveSelectedMail(
            Outlook.Application application,
            List<Outlook.MailItem> selectedMail,
            string targetEntryId,
            string targetStoreId,
            bool markAsRead,
            int skippedNonMail,
            FolderEnumerationWarnings folderWarnings)
        {
            Outlook.NameSpace session = null;
            Outlook.MAPIFolder targetFolder = null;
            try
            {
                try
                {
                    session = application.Session;
                    targetFolder = session.GetFolderFromID(targetEntryId, targetStoreId);
                }
                catch (Exception ex)
                {
                    QuickMoveLog.Write("failed to resolve target folder.", ex);
                    MessageBox.Show("The selected target folder is no longer available. It may have been moved, renamed, or deleted.", "Quick Move", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                if (targetFolder == null)
                {
                    MessageBox.Show("Select a target folder before moving mail.", "Quick Move", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                var moved = 0;
                var targetIdentity = CreateFolderIdentity(targetFolder);
                var skippedSameFolder = new List<string>();
                var failures = new List<string>();

                // One id per Quick Move action, so the Undo dialog can group and pre-check the
                // items moved together.
                var batchId = Guid.NewGuid().ToString("N");
                var undoEntries = new List<UndoEntry>();

                foreach (var mail in selectedMail)
                {
                    var subject = GetMailSubject(mail);
                    try
                    {
                        var source = CaptureSource(mail, subject);
                        if (IsSameFolder(source, targetIdentity))
                        {
                            skippedSameFolder.Add(subject);
                            continue;
                        }

                        var originalUnRead = GetUnRead(mail);
                        if (markAsRead)
                        {
                            mail.UnRead = false;
                            mail.Save();
                        }

                        var movedItem = mail.Move(targetFolder);
                        moved++;
                        try
                        {
                            var entry = new UndoEntry(
                                DateTime.UtcNow,
                                GetMovedEntryId(movedItem),
                                targetStoreId,
                                source.EntryId,
                                source.StoreId,
                                source.Path,
                                targetIdentity.DisplayPath,
                                originalUnRead,
                                subject,
                                batchId);
                            if (entry.HasUndoIdentity)
                            {
                                undoEntries.Add(entry);
                            }
                            else
                            {
                                QuickMoveLog.Write("moved '" + subject + "' but could not record it for undo (missing folder ids).");
                            }
                        }
                        finally
                        {
                            ComUtil.Release(movedItem);
                        }
                    }
                    catch (Exception ex)
                    {
                        failures.Add(subject);
                        QuickMoveLog.Write("failed to move '" + subject + "'.", ex);
                    }
                }

                if (undoEntries.Count > 0)
                {
                    UndoStore.Append(undoEntries);
                    InvalidateUndoButton();
                }

                LogSameFolderSkips(skippedSameFolder, targetIdentity);
                ShowSummaryIfNeeded(moved, skippedNonMail, skippedSameFolder, failures, folderWarnings);
            }
            finally
            {
                ComUtil.Release(targetFolder);
                ComUtil.Release(session);
            }
        }

        private static void ExecuteGoToFolder()
        {
            var application = Globals.ThisAddIn.Application;
            Outlook.Explorer explorer = null;
            try
            {
                explorer = application.ActiveExplorer();
                if (explorer == null)
                {
                    MessageBox.Show("No active Outlook explorer window is available.", "Go to Folder", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                FolderEnumerationResult folderResult;
                using (BusyCursor.Show())
                {
                    folderResult = OutlookFolderEnumerator.GetMailFolders(application);
                }

                if (folderResult.Folders.Count == 0)
                {
                    var message = folderResult.Warnings.Count > 0
                        ? "No suitable mail folders were found. Some Outlook data files could not be read. Check the Quick Move log for details."
                        : "No suitable mail folders were found.";
                    MessageBox.Show(message, "Go to Folder", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                string targetEntryId;
                string targetStoreId;
                FolderEnumerationWarnings folderWarnings;
                using (var form = new FolderPickerForm(
                    folderResult.Folders,
                    FolderPickerOptions.ForGoToFolder(),
                    folderResult.Warnings,
                    delegate { return RefreshFolderList(application); }))
                {
                    if (form.ShowDialog() != DialogResult.OK)
                    {
                        return;
                    }

                    targetEntryId = form.SelectedFolderEntryId;
                    targetStoreId = form.SelectedFolderStoreId;
                    folderWarnings = form.FolderWarnings;
                }

                if (string.IsNullOrEmpty(targetEntryId) || string.IsNullOrEmpty(targetStoreId))
                {
                    MessageBox.Show("Select a folder to go to.", "Go to Folder", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                ShowGoToFolderEnumerationWarnings(folderWarnings);
                NavigateToFolder(application, explorer, targetEntryId, targetStoreId);
            }
            finally
            {
                ComUtil.Release(explorer);
            }
        }

        private static void ExecuteGoToSelectedMailFolder()
        {
            const string title = "Go to Mail Folder";

            var application = Globals.ThisAddIn.Application;
            Outlook.Explorer explorer = null;
            Outlook.Selection selection = null;
            Outlook.MailItem selectedMail = null;

            try
            {
                explorer = application.ActiveExplorer();
                if (explorer == null)
                {
                    MessageBox.Show("No active Outlook explorer window is available.", title, MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                selection = explorer.Selection;
                if (selection == null || selection.Count == 0)
                {
                    MessageBox.Show("Select a mail item before using Go to Mail Folder.", title, MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                if (selection.Count > 1)
                {
                    var result = MessageBox.Show(
                        "Multiple items are selected. Go to Mail Folder will jump to the folder of the first mail item in the current selection. Continue?",
                        title,
                        MessageBoxButtons.OKCancel,
                        MessageBoxIcon.Warning);
                    if (result != DialogResult.OK)
                    {
                        return;
                    }
                }

                for (int i = 1; i <= selection.Count; i++)
                {
                    object selectedItem = null;
                    try
                    {
                        selectedItem = selection[i];
                    }
                    catch (Exception ex)
                    {
                        QuickMoveLog.Write("go to mail folder: failed to read selected item at index " + i + ".", ex);
                        continue;
                    }

                    selectedMail = selectedItem as Outlook.MailItem;
                    if (selectedMail != null)
                    {
                        break;
                    }

                    ComUtil.Release(selectedItem);
                }

                if (selectedMail == null)
                {
                    MessageBox.Show("The current selection does not contain a mail item.", title, MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                var subject = GetMailSubject(selectedMail);
                var source = CaptureSource(selectedMail, subject);
                if (string.IsNullOrEmpty(source.EntryId) || string.IsNullOrEmpty(source.StoreId))
                {
                    QuickMoveLog.Write("go to mail folder: could not determine source folder for '" + subject + "'.");
                    MessageBox.Show("Could not determine the selected mail item's folder. Check the Quick Move log for details.", title, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                if (IsCurrentFolder(explorer, source))
                {
                    var folderPath = string.IsNullOrWhiteSpace(source.Path) ? "(unknown folder)" : source.Path;
                    MessageBox.Show("Already in this folder:" + Environment.NewLine + folderPath, title, MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                NavigateToFolder(application, explorer, source.EntryId, source.StoreId, title, "go to mail folder");
            }
            finally
            {
                ComUtil.Release(selectedMail);
                ComUtil.Release(selection);
                ComUtil.Release(explorer);
            }
        }

        private static FolderEnumerationResult RefreshFolderList(Outlook.Application application)
        {
            // Refresh is a user-initiated, authoritative re-sync: rebuild the store-root baseline
            // from a full scan, drop the caches, then re-enumerate from the refreshed saved roots.
            // Only clean scans reconcile/prune removed stores; partial scans merge successful roots
            // and keep the existing baseline for stores that could not be read.
            var refreshResult = OutlookFolderEnumerator.RefreshStoreRoots(application);
            OutlookFolderEnumerator.InvalidateCache();
            var folderResult = OutlookFolderEnumerator.GetMailFolders(application);
            return AddRefreshWarnings(folderResult, refreshResult);
        }

        private static FolderEnumerationResult AddRefreshWarnings(FolderEnumerationResult folderResult, StoreEnumerationResult refreshResult)
        {
            if (folderResult == null || refreshResult == null || refreshResult.Errors.Count == 0)
            {
                return folderResult;
            }

            var warnings = new FolderEnumerationWarnings();
            warnings.AddRange(folderResult.Warnings);
            foreach (var error in refreshResult.Errors)
            {
                warnings.Add(FolderWarningKind.StoreUnreadable, error);
            }

            return new FolderEnumerationResult(folderResult.Folders, warnings);
        }

        private static void ShowGoToFolderEnumerationWarnings(FolderEnumerationWarnings folderWarnings)
        {
            if (folderWarnings == null || folderWarnings.Count == 0)
            {
                return;
            }

            var message = new StringBuilder();
            message.AppendLine("Some folders could not be read, so the list may be incomplete (" + folderWarnings.Count + "):");
            message.Append(BuildWarningBreakdown(folderWarnings));
            if (HasStoreUnreadableWarnings(folderWarnings))
            {
                message.AppendLine(StoreRetryHint);
            }

            message.AppendLine("Details are in the log: %TEMP%\\OutlookQuickMove.log");

            MessageBox.Show(message.ToString(), "Go to Folder", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private static void NavigateToFolder(Outlook.Application application, Outlook.Explorer explorer, string targetEntryId, string targetStoreId)
        {
            NavigateToFolder(application, explorer, targetEntryId, targetStoreId, "Go to Folder", "go to folder");
        }

        private static void NavigateToFolder(Outlook.Application application, Outlook.Explorer explorer, string targetEntryId, string targetStoreId, string title, string logPrefix)
        {
            Outlook.NameSpace session = null;
            Outlook.MAPIFolder targetFolder = null;
            try
            {
                try
                {
                    session = application.Session;
                    targetFolder = session.GetFolderFromID(targetEntryId, targetStoreId);
                }
                catch (Exception ex)
                {
                    QuickMoveLog.Write(logPrefix + ": failed to resolve target folder.", ex);
                    MessageBox.Show("The selected folder is no longer available. It may have been moved, renamed, or deleted.", title, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                if (targetFolder == null)
                {
                    MessageBox.Show("Select a folder to go to.", title, MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                try
                {
                    explorer.CurrentFolder = targetFolder;
                    explorer.Activate();
                }
                catch (Exception ex)
                {
                    QuickMoveLog.Write(logPrefix + ": failed to navigate to the selected folder.", ex);
                    MessageBox.Show("Could not switch to the selected folder. Check the Quick Move log for details.", title, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            finally
            {
                ComUtil.Release(targetFolder);
                ComUtil.Release(session);
            }
        }

        private static void ExecuteSettings()
        {
            var application = Globals.ThisAddIn.Application;
            StoreEnumerationResult storeResult;
            using (BusyCursor.Show())
            {
                storeResult = OutlookFolderEnumerator.GetStores(application);
            }

            if (storeResult.Stores.Count == 0)
            {
                MessageBox.Show("No Outlook data files were found.", "Quick Move", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            bool saved;
            using (var form = new SettingsForm(
                storeResult.Stores,
                StoreFilterSettings.LoadEnabledStoreKeys(),
                FrequentTargetStore.LoadHistory(),
                FrequentTargetStore.GetMaxCount(),
                UndoStore.GetMaxCount()))
            {
                if (form.ShowDialog() != DialogResult.OK)
                {
                    return;
                }

                // When every data file is selected, persist an empty filter. An empty filter
                // means "no filter" (search all stores), which keeps any newly added data file
                // included by default rather than silently excluded.
                var selectedStoreKeys = form.SelectedStoreKeys.Count == storeResult.Stores.Count
                    ? new List<string>()
                    : form.SelectedStoreKeys;

                // Save the new cap first (it trims the stored list), then replace the list with
                // whatever the user left after any deletions.
                var storeSaved = StoreFilterSettings.SaveEnabledStoreKeys(selectedStoreKeys);

                // Persist root identities for ALL known stores (decoupled from the enable filter),
                // so even "all selected" gets the saved-root fast path. Only clean scans replace the
                // baseline and prune removed stores; partial scans merge successes and preserve roots
                // for stores that could not be read.
                var storeRootSaved = storeResult.Errors.Count == 0
                    ? StoreFilterSettings.SaveEnabledStoreEntries(storeResult.Stores)
                    : StoreFilterSettings.MergeStoreRoots(storeResult.Stores);
                var capSaved = FrequentTargetStore.SaveMaxCount(form.MaxFrequentCount);
                var frequentSaved = FrequentTargetStore.ReplaceAll(form.FrequentTargets);
                var undoCapSaved = UndoStore.SaveMaxCount(form.MaxUndoCount);
                var undoCleared = !form.ClearUndoHistoryRequested || UndoStore.Save(new List<UndoEntry>());
                saved = storeSaved && storeRootSaved && capSaved && frequentSaved && undoCapSaved && undoCleared;
            }

            // The undo cap or a clear may have changed whether anything is undoable.
            InvalidateUndoButton();

            // The data-file selection may have changed, so drop the cached folder list and
            // re-enumerate on the next Quick Move / Go to Folder.
            OutlookFolderEnumerator.InvalidateCache();

            if (!saved)
            {
                MessageBox.Show("Your selection could not be saved. Check the Quick Move log for details.", "Quick Move", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (storeResult.Errors.Count > 0)
            {
                MessageBox.Show(
                    "Settings saved. Some Outlook data files could not be read; their existing entries were kept. "
                    + "If Outlook is busy, reopen Settings in a minute to finish syncing.",
                    "Quick Move",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
        }

        private static void ExecuteUndo()
        {
            var history = UndoStore.LoadAll();
            if (history.Count == 0)
            {
                MessageBox.Show("There are no Quick Move actions to undo.", "Quick Move", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            List<UndoEntry> selected;
            bool clearAll;
            using (var form = new UndoHistoryForm(history))
            {
                if (form.ShowDialog() != DialogResult.OK)
                {
                    return;
                }

                selected = form.SelectedEntries;
                clearAll = form.ClearAllRequested;
            }

            if (clearAll)
            {
                UndoStore.Save(new List<UndoEntry>());
                InvalidateUndoButton();
                return;
            }

            if (selected == null || selected.Count == 0)
            {
                return;
            }

            var application = Globals.ThisAddIn.Application;
            using (BusyCursor.Show())
            {
                RestoreEntries(application, history, selected);
            }
        }

        private static void RestoreEntries(Outlook.Application application, List<UndoEntry> history, List<UndoEntry> selected)
        {
            Outlook.NameSpace session = null;
            var restored = 0;
            var notFound = 0;
            var failed = 0;

            // Keys to drop from history: successfully undone, plus entries whose item no longer
            // exists (they can never be undone, so keeping them just clutters the list). Entries
            // that failed for other reasons are kept so the user can retry.
            var keysToDrop = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            try
            {
                session = application.Session;
                foreach (var entry in selected)
                {
                    switch (RestoreOne(session, entry))
                    {
                        case RestoreOutcome.Restored:
                            restored++;
                            keysToDrop.Add(entry.Key);
                            break;
                        case RestoreOutcome.NotFound:
                            notFound++;
                            keysToDrop.Add(entry.Key);
                            break;
                        default:
                            failed++;
                            break;
                    }
                }
            }
            finally
            {
                ComUtil.Release(session);
            }

            var remaining = history.Where(e => !keysToDrop.Contains(e.Key)).ToList();
            UndoStore.Save(remaining);
            InvalidateUndoButton();

            ShowUndoSummary(restored, notFound, failed);
        }

        private static RestoreOutcome RestoreOne(Outlook.NameSpace session, UndoEntry entry)
        {
            object item = null;
            Outlook.MAPIFolder sourceFolder = null;
            object movedBack = null;
            try
            {
                try
                {
                    item = session.GetItemFromID(entry.MovedEntryId, entry.MovedStoreId);
                }
                catch (Exception ex)
                {
                    QuickMoveLog.Write("undo: moved item no longer found for '" + entry.Subject + "'.", ex);
                    return RestoreOutcome.NotFound;
                }

                var mail = item as Outlook.MailItem;
                if (mail == null)
                {
                    QuickMoveLog.Write("undo: moved item is not a mail item for '" + entry.Subject + "'.");
                    return RestoreOutcome.NotFound;
                }

                try
                {
                    sourceFolder = session.GetFolderFromID(entry.SourceFolderEntryId, entry.SourceFolderStoreId);
                }
                catch (Exception ex)
                {
                    QuickMoveLog.Write("undo: original folder no longer available for '" + entry.Subject + "'.", ex);
                    return RestoreOutcome.Failed;
                }

                if (sourceFolder == null)
                {
                    return RestoreOutcome.Failed;
                }

                movedBack = mail.Move(sourceFolder);
                var restoredMail = movedBack as Outlook.MailItem;
                if (restoredMail != null)
                {
                    try
                    {
                        restoredMail.UnRead = entry.OriginalUnRead;
                        restoredMail.Save();
                    }
                    catch (Exception ex)
                    {
                        QuickMoveLog.Write("undo: failed to restore read state for '" + entry.Subject + "'.", ex);
                    }
                }

                return RestoreOutcome.Restored;
            }
            catch (Exception ex)
            {
                QuickMoveLog.Write("undo: failed to move '" + entry.Subject + "' back.", ex);
                return RestoreOutcome.Failed;
            }
            finally
            {
                // 'mail'/'restoredMail' alias 'item'/'movedBack', so release only the originals to
                // avoid over-releasing the same RCW.
                ComUtil.Release(movedBack);
                ComUtil.Release(sourceFolder);
                ComUtil.Release(item);
            }
        }

        private static void ShowUndoSummary(int restored, int notFound, int failed)
        {
            if (notFound == 0 && failed == 0)
            {
                MessageBox.Show("Moved " + restored + " item(s) back to the original folder(s).", "Quick Move", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var message = new StringBuilder();
            message.AppendLine("Undo summary:");
            message.AppendLine("Restored: " + restored);
            if (notFound > 0)
            {
                message.AppendLine("No longer found (removed from history): " + notFound);
            }

            if (failed > 0)
            {
                message.AppendLine("Could not be restored (kept in history): " + failed);
            }

            MessageBox.Show(message.ToString(), "Quick Move", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        private static void ShowSummaryIfNeeded(int moved, int skippedNonMail, List<string> skippedSameFolder, List<string> failures, FolderEnumerationWarnings folderWarnings)
        {
            var warningCount = folderWarnings == null ? 0 : folderWarnings.Count;
            if (skippedNonMail == 0 && skippedSameFolder.Count == 0 && failures.Count == 0 && warningCount == 0)
            {
                return;
            }

            var message = new StringBuilder();
            message.AppendLine("Quick Move summary:");
            message.AppendLine("Moved: " + moved);

            if (skippedNonMail > 0)
            {
                message.AppendLine("Skipped non-mail items: " + skippedNonMail);
            }

            if (skippedSameFolder.Count > 0)
            {
                message.AppendLine("Skipped because source and target folder are the same: " + skippedSameFolder.Count);
                foreach (var subject in skippedSameFolder.GetRange(0, Math.Min(5, skippedSameFolder.Count)))
                {
                    message.AppendLine("- " + subject);
                }
            }

            if (failures.Count > 0)
            {
                message.AppendLine("Failed to move: " + failures.Count);
                foreach (var subject in failures.GetRange(0, Math.Min(5, failures.Count)))
                {
                    message.AppendLine("- " + subject);
                }
            }

            if (warningCount > 0)
            {
                message.AppendLine("Folder enumeration warnings: " + warningCount);
                message.Append(BuildWarningBreakdown(folderWarnings));
                if (HasStoreUnreadableWarnings(folderWarnings))
                {
                    message.AppendLine(StoreRetryHint);
                }

                message.AppendLine("Details are in the log: %TEMP%\\OutlookQuickMove.log");
            }

            MessageBox.Show(message.ToString(), "Quick Move", MessageBoxButtons.OK, failures.Count > 0 || skippedSameFolder.Count > 0 ? MessageBoxIcon.Warning : MessageBoxIcon.Information);
        }

        // Shown when one or more data files could not be read. Rebuilding the list binds every data
        // file, which can transiently fail when Outlook is busy (e.g. still mounting stores at
        // startup), so a retry often succeeds.
        private const string StoreRetryHint =
            "Tip: if Outlook is busy, press Refresh in the dialog (or reopen Settings) again in a minute.";

        private static bool HasStoreUnreadableWarnings(FolderEnumerationWarnings folderWarnings)
        {
            int count;
            return folderWarnings != null
                && folderWarnings.Counts.TryGetValue(FolderWarningKind.StoreUnreadable, out count)
                && count > 0;
        }

        /// <summary>
        /// Renders folder-enumeration warning counts grouped by cause (one indented line per
        /// category), so the user sees what kind of folders were skipped rather than a bare total.
        /// </summary>
        private static string BuildWarningBreakdown(FolderEnumerationWarnings folderWarnings)
        {
            var breakdown = new StringBuilder();
            if (folderWarnings == null)
            {
                return breakdown.ToString();
            }

            foreach (var pair in folderWarnings.Counts.OrderByDescending(p => p.Value))
            {
                breakdown.AppendLine("- " + FolderEnumerationWarnings.Describe(pair.Key) + ": " + pair.Value);
            }

            return breakdown.ToString();
        }

        private static string GetMailSubject(Outlook.MailItem mail)
        {
            try
            {
                return string.IsNullOrWhiteSpace(mail.Subject) ? "(no subject)" : mail.Subject;
            }
            catch (Exception ex)
            {
                QuickMoveLog.Write("failed to read mail subject.", ex);
                return "(unavailable subject)";
            }
        }

        /// <summary>
        /// Reads the mail's current (source) folder identity once, before the move, for both the
        /// same-folder check and the undo record. Never throws; missing fields come back empty.
        /// </summary>
        private static SourceInfo CaptureSource(Outlook.MailItem mail, string subject)
        {
            Outlook.MAPIFolder sourceFolder = null;
            try
            {
                sourceFolder = mail.Parent as Outlook.MAPIFolder;
                if (sourceFolder == null)
                {
                    return new SourceInfo(string.Empty, string.Empty, string.Empty);
                }

                return new SourceInfo(
                    GetFolderEntryId(sourceFolder),
                    GetFolderStoreId(sourceFolder),
                    GetFolderPath(sourceFolder));
            }
            catch (Exception ex)
            {
                QuickMoveLog.Write("failed to resolve source folder for '" + subject + "'.", ex);
                return new SourceInfo(string.Empty, string.Empty, string.Empty);
            }
            finally
            {
                ComUtil.Release(sourceFolder);
            }
        }

        private static bool IsSameFolder(SourceInfo source, FolderIdentity targetIdentity)
        {
            if (source == null || targetIdentity == null)
            {
                return false;
            }

            if (!string.IsNullOrEmpty(source.EntryId) && !string.IsNullOrEmpty(targetIdentity.EntryId))
            {
                return string.Equals(source.EntryId, targetIdentity.EntryId, StringComparison.OrdinalIgnoreCase);
            }

            return !string.IsNullOrEmpty(source.NormalizedPath)
                && !string.IsNullOrEmpty(targetIdentity.NormalizedPath)
                && string.Equals(source.NormalizedPath, targetIdentity.NormalizedPath, StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsSameFolder(SourceInfo first, SourceInfo second)
        {
            if (first == null || second == null)
            {
                return false;
            }

            if (!string.IsNullOrEmpty(first.EntryId) && !string.IsNullOrEmpty(second.EntryId)
                && !string.IsNullOrEmpty(first.StoreId) && !string.IsNullOrEmpty(second.StoreId))
            {
                return string.Equals(first.EntryId, second.EntryId, StringComparison.OrdinalIgnoreCase)
                    && string.Equals(first.StoreId, second.StoreId, StringComparison.OrdinalIgnoreCase);
            }

            if (!string.IsNullOrEmpty(first.EntryId) && !string.IsNullOrEmpty(second.EntryId))
            {
                return string.Equals(first.EntryId, second.EntryId, StringComparison.OrdinalIgnoreCase);
            }

            return !string.IsNullOrEmpty(first.NormalizedPath)
                && !string.IsNullOrEmpty(second.NormalizedPath)
                && string.Equals(first.NormalizedPath, second.NormalizedPath, StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsCurrentFolder(Outlook.Explorer explorer, SourceInfo target)
        {
            Outlook.MAPIFolder currentFolder = null;
            try
            {
                currentFolder = explorer.CurrentFolder as Outlook.MAPIFolder;
                if (currentFolder == null)
                {
                    return false;
                }

                var current = new SourceInfo(
                    GetFolderEntryId(currentFolder),
                    GetFolderStoreId(currentFolder),
                    GetFolderPath(currentFolder));
                return IsSameFolder(target, current);
            }
            catch (Exception ex)
            {
                QuickMoveLog.Write("go to mail folder: failed to compare current folder.", ex);
                return false;
            }
            finally
            {
                ComUtil.Release(currentFolder);
            }
        }

        private static string GetFolderStoreId(Outlook.MAPIFolder folder)
        {
            try
            {
                return folder.StoreID;
            }
            catch (Exception ex)
            {
                QuickMoveLog.Write("failed to read folder store id.", ex);
                return string.Empty;
            }
        }

        private static bool GetUnRead(Outlook.MailItem mail)
        {
            try
            {
                return mail.UnRead;
            }
            catch (Exception ex)
            {
                QuickMoveLog.Write("failed to read mail read state.", ex);
                return false;
            }
        }

        private static string GetMovedEntryId(object movedItem)
        {
            try
            {
                var mail = movedItem as Outlook.MailItem;
                return mail == null ? string.Empty : mail.EntryID;
            }
            catch (Exception ex)
            {
                QuickMoveLog.Write("failed to read moved item entry id.", ex);
                return string.Empty;
            }
        }

        private static FolderIdentity CreateFolderIdentity(Outlook.MAPIFolder folder)
        {
            var folderPath = GetFolderPath(folder);
            return new FolderIdentity(
                GetFolderEntryId(folder),
                NormalizeFolderPath(folderPath),
                string.IsNullOrWhiteSpace(folderPath) ? "(unknown target folder)" : folderPath);
        }

        private static void LogSameFolderSkips(List<string> skippedSameFolder, FolderIdentity targetIdentity)
        {
            if (skippedSameFolder.Count == 0)
            {
                return;
            }

            var sampleSubjects = skippedSameFolder.GetRange(0, Math.Min(5, skippedSameFolder.Count));
            var suffix = skippedSameFolder.Count > sampleSubjects.Count ? "; ..." : string.Empty;
            QuickMoveLog.Write("skipped " + skippedSameFolder.Count + " item(s) because source and target folders are the same: "
                + targetIdentity.DisplayPath + ". Subjects: " + string.Join("; ", sampleSubjects) + suffix);
        }

        private static string GetFolderEntryId(Outlook.MAPIFolder folder)
        {
            try
            {
                return folder.EntryID;
            }
            catch (Exception ex)
            {
                QuickMoveLog.Write("failed to read folder entry id.", ex);
                return string.Empty;
            }
        }

        private static string GetFolderPath(Outlook.MAPIFolder folder)
        {
            try
            {
                return folder.FolderPath;
            }
            catch (Exception ex)
            {
                QuickMoveLog.Write("failed to read folder path.", ex);
                return string.Empty;
            }
        }

        private static string NormalizeFolderPath(string folderPath)
        {
            return string.IsNullOrWhiteSpace(folderPath)
                ? string.Empty
                : folderPath.Trim().TrimStart('\\');
        }

        private enum RestoreOutcome
        {
            Restored,
            NotFound,
            Failed
        }

        private sealed class SourceInfo
        {
            public SourceInfo(string entryId, string storeId, string path)
            {
                EntryId = entryId ?? string.Empty;
                StoreId = storeId ?? string.Empty;
                Path = path ?? string.Empty;
                NormalizedPath = NormalizeFolderPath(Path);
            }

            public string EntryId { get; }

            public string StoreId { get; }

            public string Path { get; }

            public string NormalizedPath { get; }
        }

        private sealed class FolderIdentity
        {
            public FolderIdentity(string entryId, string normalizedPath, string displayPath)
            {
                EntryId = entryId ?? string.Empty;
                NormalizedPath = normalizedPath ?? string.Empty;
                DisplayPath = displayPath ?? string.Empty;
            }

            public string EntryId { get; }

            public string NormalizedPath { get; }

            public string DisplayPath { get; }
        }
    }
}
