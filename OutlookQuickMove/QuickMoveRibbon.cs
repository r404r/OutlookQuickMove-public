using System;
using System.Collections.Generic;
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
        public string GetCustomUI(string ribbonId)
        {
            ThisAddIn.DebugLog("GetCustomUI ribbonId=" + (ribbonId ?? "(null)"));
            if (!string.Equals(ribbonId, "Microsoft.Outlook.Explorer", StringComparison.OrdinalIgnoreCase))
            {
                return string.Empty;
            }

            return @"<?xml version=""1.0"" encoding=""UTF-8""?>
<customUI xmlns=""http://schemas.microsoft.com/office/2009/07/customui"">
  <ribbon>
    <tabs>
      <tab id=""OutlookQuickMoveTab"" label=""Quick Move"" insertAfterMso=""TabMail"">
        <group id=""OutlookQuickMoveActionsGroup"" label=""Actions"">
          <button id=""QuickMoveButton""
                  label=""Quick Move""
                  size=""large""
                  imageMso=""MoveToFolder""
                  onAction=""OnQuickMove"" />
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
                using (var form = new MoveToFolderForm(folderResult.Folders))
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

                MoveSelectedMail(application, selectedMail, targetEntryId, targetStoreId, markAsRead, skippedNonMail, folderResult.Errors);
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
            List<string> folderErrors)
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

                foreach (var mail in selectedMail)
                {
                    var subject = GetMailSubject(mail);
                    try
                    {
                        if (IsMailAlreadyInTargetFolder(mail, targetIdentity, subject))
                        {
                            skippedSameFolder.Add(subject);
                            continue;
                        }

                        if (markAsRead)
                        {
                            mail.UnRead = false;
                            mail.Save();
                        }

                        var movedItem = mail.Move(targetFolder);
                        moved++;
                        ComUtil.Release(movedItem);
                    }
                    catch (Exception ex)
                    {
                        failures.Add(subject);
                        QuickMoveLog.Write("failed to move '" + subject + "'.", ex);
                    }
                }

                LogSameFolderSkips(skippedSameFolder, targetIdentity);
                ShowSummaryIfNeeded(moved, skippedNonMail, skippedSameFolder, failures, folderErrors);
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
                FrequentTargetStore.GetMaxCount()))
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
                saved = storeSaved && capSaved && frequentSaved;
            }

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

        private static void ShowSummaryIfNeeded(int moved, int skippedNonMail, List<string> skippedSameFolder, List<string> failures, List<string> folderErrors)
        {
            if (skippedNonMail == 0 && skippedSameFolder.Count == 0 && failures.Count == 0 && folderErrors.Count == 0)
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

            if (folderErrors.Count > 0)
            {
                message.AppendLine("Folder enumeration warnings: " + folderErrors.Count);
            }

            MessageBox.Show(message.ToString(), "Quick Move", MessageBoxButtons.OK, failures.Count > 0 || skippedSameFolder.Count > 0 ? MessageBoxIcon.Warning : MessageBoxIcon.Information);
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

        private static bool IsMailAlreadyInTargetFolder(Outlook.MailItem mail, FolderIdentity targetIdentity, string subject)
        {
            Outlook.MAPIFolder sourceFolder = null;
            try
            {
                sourceFolder = mail.Parent as Outlook.MAPIFolder;
                return IsSameFolder(sourceFolder, targetIdentity);
            }
            catch (Exception ex)
            {
                QuickMoveLog.Write("failed to resolve source folder for '" + subject + "'.", ex);
                return false;
            }
            finally
            {
                ComUtil.Release(sourceFolder);
            }
        }

        private static bool IsSameFolder(Outlook.MAPIFolder sourceFolder, FolderIdentity targetIdentity)
        {
            if (sourceFolder == null || targetIdentity == null)
            {
                return false;
            }

            var sourceEntryId = GetFolderEntryId(sourceFolder);
            if (!string.IsNullOrEmpty(sourceEntryId) && !string.IsNullOrEmpty(targetIdentity.EntryId))
            {
                return string.Equals(sourceEntryId, targetIdentity.EntryId, StringComparison.OrdinalIgnoreCase);
            }

            var sourcePath = NormalizeFolderPath(GetFolderPath(sourceFolder));
            return !string.IsNullOrEmpty(sourcePath)
                && !string.IsNullOrEmpty(targetIdentity.NormalizedPath)
                && string.Equals(sourcePath, targetIdentity.NormalizedPath, StringComparison.OrdinalIgnoreCase);
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
