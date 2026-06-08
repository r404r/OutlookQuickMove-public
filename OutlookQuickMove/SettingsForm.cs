using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace OutlookQuickMove
{
    /// <summary>
    /// Quick Move settings, organized into tabs:
    /// - Data Files: which Outlook stores are searched for target folders.
    /// - Frequent Folders: the cap on remembered targets, plus the list with delete actions.
    /// - Undo History: the cap on remembered moves, plus a clear option.
    /// The form only collects choices; persistence is performed by the caller.
    /// </summary>
    internal sealed class SettingsForm : Form
    {
        private readonly List<StoreCandidate> stores;
        private readonly CheckedListBox listStores;
        private readonly NumericUpDown numericMaxFrequent;
        private readonly ListBox listFrequent;
        private readonly NumericUpDown numericMaxUndo;
        private readonly CheckBox checkClearUndo;

        public SettingsForm(
            IEnumerable<StoreCandidate> stores,
            HashSet<string> enabledStoreKeys,
            IEnumerable<FrequentTarget> frequentTargets,
            int maxFrequentCount,
            int maxUndoCount)
        {
            this.stores = stores == null ? new List<StoreCandidate>() : stores.ToList();

            Text = "Quick Move Settings";
            StartPosition = FormStartPosition.CenterScreen;
            MinimumSize = new Size(560, 400);
            Size = new Size(720, 480);
            Font = SystemFonts.MessageBoxFont;
            QuickMoveIcon.ApplyTo(this);

            var layout = new TableLayoutPanel
            {
                ColumnCount = 1,
                Dock = DockStyle.Fill,
                Padding = new Padding(12),
                RowCount = 2
            };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            var tabs = new TabControl { Dock = DockStyle.Fill };

            listStores = new CheckedListBox
            {
                CheckOnClick = true,
                DisplayMember = nameof(StoreCandidate.DisplayText),
                Dock = DockStyle.Fill,
                IntegralHeight = false
            };

            numericMaxFrequent = new NumericUpDown
            {
                Minimum = 0,
                Maximum = FrequentTargetStore.MaxAllowedCount,
                Value = Clamp(maxFrequentCount),
                Width = 64
            };

            listFrequent = new ListBox
            {
                Dock = DockStyle.Fill,
                IntegralHeight = false,
                SelectionMode = SelectionMode.MultiExtended
            };

            numericMaxUndo = new NumericUpDown
            {
                Minimum = 0,
                Maximum = UndoStore.MaxAllowedCount,
                Value = ClampUndo(maxUndoCount),
                Width = 64
            };

            checkClearUndo = new CheckBox
            {
                AutoSize = true,
                Text = "Clear all remembered undo history when I save"
            };

            tabs.TabPages.Add(BuildDataFilesTab(enabledStoreKeys));
            tabs.TabPages.Add(BuildFrequentTab(frequentTargets));
            tabs.TabPages.Add(BuildUndoTab());

            var buttonOk = new Button
            {
                DialogResult = DialogResult.None,
                Margin = new Padding(8, 0, 0, 0),
                Size = new Size(88, 28),
                Text = "OK"
            };
            buttonOk.Click += delegate { ConfirmSettings(); };

            var buttonCancel = new Button
            {
                DialogResult = DialogResult.Cancel,
                Margin = new Padding(8, 0, 0, 0),
                Size = new Size(88, 28),
                Text = "Cancel"
            };

            var buttonPanel = new FlowLayoutPanel
            {
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.RightToLeft,
                Margin = new Padding(0, 8, 0, 0),
                Padding = new Padding(0)
            };
            buttonPanel.Controls.Add(buttonCancel);
            buttonPanel.Controls.Add(buttonOk);

            layout.Controls.Add(tabs, 0, 0);
            layout.Controls.Add(buttonPanel, 0, 1);
            Controls.Add(layout);

            AcceptButton = buttonOk;
            CancelButton = buttonCancel;
        }

        public List<string> SelectedStoreKeys { get; private set; }

        public List<StoreCandidate> SelectedStores { get; private set; }

        public int MaxFrequentCount { get; private set; }

        public List<FrequentTarget> FrequentTargets { get; private set; }

        public int MaxUndoCount { get; private set; }

        public bool ClearUndoHistoryRequested { get; private set; }

        private TabPage BuildDataFilesTab(HashSet<string> enabledStoreKeys)
        {
            var page = new TabPage("Data Files") { Padding = new Padding(8), UseVisualStyleBackColor = true };

            var layout = new TableLayoutPanel { ColumnCount = 1, Dock = DockStyle.Fill, RowCount = 3 };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            var label = new Label
            {
                AutoSize = true,
                Margin = new Padding(0, 0, 0, 6),
                Text = "Search for target folders in these Outlook data files:"
            };

            var hasSavedFilter = enabledStoreKeys != null && enabledStoreKeys.Count > 0;
            listStores.BeginUpdate();
            try
            {
                listStores.Items.Clear();
                foreach (var store in stores)
                {
                    var isChecked = !hasSavedFilter || enabledStoreKeys.Contains(store.StoreKey);
                    listStores.Items.Add(store, isChecked);
                }
            }
            finally
            {
                listStores.EndUpdate();
            }

            var buttonSelectAll = new Button { AutoSize = true, Margin = new Padding(0, 6, 8, 0), Text = "Select All" };
            buttonSelectAll.Click += delegate { SetAllStoresChecked(true); };

            var buttonClear = new Button { AutoSize = true, Margin = new Padding(0, 6, 0, 0), Text = "Clear" };
            buttonClear.Click += delegate { SetAllStoresChecked(false); };

            var buttons = new FlowLayoutPanel
            {
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.LeftToRight,
                Margin = new Padding(0)
            };
            buttons.Controls.Add(buttonSelectAll);
            buttons.Controls.Add(buttonClear);

            layout.Controls.Add(label, 0, 0);
            layout.Controls.Add(listStores, 0, 1);
            layout.Controls.Add(buttons, 0, 2);
            page.Controls.Add(layout);
            return page;
        }

        private TabPage BuildFrequentTab(IEnumerable<FrequentTarget> frequentTargets)
        {
            var page = new TabPage("Frequent Folders") { Padding = new Padding(8), UseVisualStyleBackColor = true };

            var layout = new TableLayoutPanel { ColumnCount = 1, Dock = DockStyle.Fill, RowCount = 3 };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            var capLabel = new Label { Anchor = AnchorStyles.Left, AutoSize = true, Margin = new Padding(0, 6, 6, 0), Text = "Show at most" };
            var foldersLabel = new Label { Anchor = AnchorStyles.Left, AutoSize = true, Margin = new Padding(6, 6, 0, 0), Text = "folders at the top (0 = off)" };
            numericMaxFrequent.Margin = new Padding(0, 4, 0, 0);

            var capRow = new FlowLayoutPanel
            {
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.LeftToRight,
                Margin = new Padding(0, 0, 0, 6)
            };
            capRow.Controls.Add(capLabel);
            capRow.Controls.Add(numericMaxFrequent);
            capRow.Controls.Add(foldersLabel);

            listFrequent.BeginUpdate();
            try
            {
                listFrequent.Items.Clear();
                if (frequentTargets != null)
                {
                    foreach (var target in frequentTargets)
                    {
                        listFrequent.Items.Add(target);
                    }
                }
            }
            finally
            {
                listFrequent.EndUpdate();
            }

            var buttonDelete = new Button { AutoSize = true, Margin = new Padding(0, 6, 8, 0), Text = "Delete Selected" };
            buttonDelete.Click += delegate { DeleteSelectedFrequent(); };

            var buttonClearAll = new Button { AutoSize = true, Margin = new Padding(0, 6, 0, 0), Text = "Clear All" };
            buttonClearAll.Click += delegate { listFrequent.Items.Clear(); };

            var buttons = new FlowLayoutPanel
            {
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.LeftToRight,
                Margin = new Padding(0)
            };
            buttons.Controls.Add(buttonDelete);
            buttons.Controls.Add(buttonClearAll);

            layout.Controls.Add(capRow, 0, 0);
            layout.Controls.Add(listFrequent, 0, 1);
            layout.Controls.Add(buttons, 0, 2);
            page.Controls.Add(layout);
            return page;
        }

        private TabPage BuildUndoTab()
        {
            var page = new TabPage("Undo History") { Padding = new Padding(8), UseVisualStyleBackColor = true };

            var layout = new TableLayoutPanel { ColumnCount = 1, Dock = DockStyle.Fill, RowCount = 3 };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            var capLabel = new Label { Anchor = AnchorStyles.Left, AutoSize = true, Margin = new Padding(0, 6, 6, 0), Text = "Remember the last" };
            var movesLabel = new Label { Anchor = AnchorStyles.Left, AutoSize = true, Margin = new Padding(6, 6, 0, 0), Text = "moves for undo (0 = off)" };
            numericMaxUndo.Margin = new Padding(0, 4, 0, 0);

            var capRow = new FlowLayoutPanel
            {
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.LeftToRight,
                Margin = new Padding(0, 0, 0, 6)
            };
            capRow.Controls.Add(capLabel);
            capRow.Controls.Add(numericMaxUndo);
            capRow.Controls.Add(movesLabel);

            var hint = new Label
            {
                AutoSize = true,
                ForeColor = SystemColors.GrayText,
                Margin = new Padding(0, 0, 0, 6),
                Text = "Use the \"Undo Quick Move\" button on the ribbon to move recent items back."
            };

            checkClearUndo.Margin = new Padding(0, 0, 0, 0);

            layout.Controls.Add(capRow, 0, 0);
            layout.Controls.Add(hint, 0, 1);
            layout.Controls.Add(checkClearUndo, 0, 2);
            page.Controls.Add(layout);
            return page;
        }

        private void SetAllStoresChecked(bool isChecked)
        {
            for (int i = 0; i < listStores.Items.Count; i++)
            {
                listStores.SetItemChecked(i, isChecked);
            }
        }

        private void DeleteSelectedFrequent()
        {
            var selected = listFrequent.SelectedItems.Cast<object>().ToList();
            foreach (var item in selected)
            {
                listFrequent.Items.Remove(item);
            }
        }

        private void ConfirmSettings()
        {
            var selectedKeys = listStores.CheckedItems
                .Cast<StoreCandidate>()
                .Select(store => store.StoreKey)
                .ToList();

            if (selectedKeys.Count == 0)
            {
                MessageBox.Show("Select at least one data file.", "Quick Move", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            SelectedStoreKeys = selectedKeys;
            SelectedStores = listStores.CheckedItems.Cast<StoreCandidate>().ToList();
            MaxFrequentCount = (int)numericMaxFrequent.Value;
            FrequentTargets = listFrequent.Items.Cast<FrequentTarget>().ToList();
            MaxUndoCount = (int)numericMaxUndo.Value;
            ClearUndoHistoryRequested = checkClearUndo.Checked;
            DialogResult = DialogResult.OK;
            Close();
        }

        private static decimal Clamp(int value)
        {
            if (value < 0)
            {
                return 0;
            }

            return value > FrequentTargetStore.MaxAllowedCount ? FrequentTargetStore.MaxAllowedCount : value;
        }

        private static decimal ClampUndo(int value)
        {
            if (value < 0)
            {
                return 0;
            }

            return value > UndoStore.MaxAllowedCount ? UndoStore.MaxAllowedCount : value;
        }
    }
}
