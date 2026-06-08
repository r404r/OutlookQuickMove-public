using System;
using System.Collections.Generic;
using System.Linq;
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
<customUI xmlns=""http://schemas.microsoft.com/office/2009/07/customui"" onLoad=""OnRibbonLoad"">
  <ribbon>
    <tabs>
      <tab id=""OutlookQuickMoveTab"" label=""Quick Move"" insertAfterMso=""TabMail"">
        <group id=""OutlookQuickMoveActionsGroup"" label=""Actions"">
          <button id=""QuickMoveButton""
                  label=""Quick Move""
                  size=""large""
                  imageMso=""MoveToFolder""
                  onAction=""OnQuickMove"" />
          <button id=""QuickJumpButton""
                  label=""Go to Folder""
                  size=""large""
                  imageMso=""FolderOpen""
                  onAction=""OnGoToFolder"" />
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
                using (var form = new FolderPickerForm(folderResult.Folders, FolderPickerOptions.ForQuickMove()))
                {
                    if (form.ShowDialog() != DialogResult.OK)
                    {
                        return;
                    }

                    targetEntryId = form.SelectedFolderEntryId;
                    targetStoreId = form.SelectedFolderStoreId;
                    markAsRead = form.MarkAsReadBeforeMoving;
                }

                if (string.IsNullOrEmpty(targetEntryId) || string.IsNullOrEmpty(targetStoreId))
                {
                    MessageBox.Show("Select a target folder before moving mail.", "Quick Move", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                MoveSelectedMail(application, selectedMail, targetEntryId, targetStoreId, markAsRead, skippedNonMail, folderResult.Warnings);
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

                ShowGoToFolderEnumerationWarnings(folderResult.Warnings);

                string targetEntryId;
                string targetStoreId;
                using (var form = new FolderPickerForm(folderResult.Folders, FolderPickerOptions.ForGoToFolder()))
                {
                    if (form.ShowDialog() != DialogResult.OK)
                    {
                        return;
                    }

                    targetEntryId = form.SelectedFolderEntryId;
                    targetStoreId = form.SelectedFolderStoreId;
                }

                if (string.IsNullOrEmpty(targetEntryId) || string.IsNullOrEmpty(targetStoreId))
                {
                    MessageBox.Show("Select a folder to go to.", "Go to Folder", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                NavigateToFolder(application, explorer, targetEntryId, targetStoreId);
            }
            finally
            {
                ComUtil.Release(explorer);
            }
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
            message.AppendLine("Details are in the log: %TEMP%\\OutlookQuickMove.log");

            MessageBox.Show(message.ToString(), "Go to Folder", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private static void NavigateToFolder(Outlook.Application application, Outlook.Explorer explorer, string targetEntryId, string targetStoreId)
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
                    QuickMoveLog.Write("go to folder: failed to resolve target folder.", ex);
                    MessageBox.Show("The selected folder is no longer available. It may have been moved, renamed, or deleted.", "Go to Folder", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                if (targetFolder == null)
                {
                    MessageBox.Show("Select a folder to go to.", "Go to Folder", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                try
                {
                    explorer.CurrentFolder = targetFolder;
                    explorer.Activate();
                }
                catch (Exception ex)
                {
                    QuickMoveLog.Write("go to folder: failed to navigate to the selected folder.", ex);
                    MessageBox.Show("Could not switch to the selected folder. Check the Quick Move log for details.", "Go to Folder", MessageBoxButtons.OK, MessageBoxIcon.Warning);
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
                var capSaved = FrequentTargetStore.SaveMaxCount(form.MaxFrequentCount);
                var frequentSaved = FrequentTargetStore.ReplaceAll(form.FrequentTargets);
                var undoCapSaved = UndoStore.SaveMaxCount(form.MaxUndoCount);
                var undoCleared = !form.ClearUndoHistoryRequested || UndoStore.Save(new List<UndoEntry>());
                saved = storeSaved && capSaved && frequentSaved && undoCapSaved && undoCleared;
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
                MessageBox.Show("Settings saved. Some Outlook data files could not be read.", "Quick Move", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
                message.AppendLine("Details are in the log: %TEMP%\\OutlookQuickMove.log");
            }

            MessageBox.Show(message.ToString(), "Quick Move", MessageBoxButtons.OK, failures.Count > 0 || skippedSameFolder.Count > 0 ? MessageBoxIcon.Warning : MessageBoxIcon.Information);
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
