using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace OutlookQuickMove
{
    internal sealed class MoveToFolderForm : Form
    {
        private readonly List<FolderCandidate> allFolders;
        private readonly Dictionary<string, FrequentTarget> frequentByKey;
        private readonly TextBox textSearch;
        private readonly ListBox listFolders;
        private readonly CheckBox checkMarkAsRead;
        private readonly Button buttonOk;

        public MoveToFolderForm(IEnumerable<FolderCandidate> folders)
        {
            allFolders = folders == null ? new List<FolderCandidate>() : folders.ToList();

            frequentByKey = new Dictionary<string, FrequentTarget>(StringComparer.OrdinalIgnoreCase);
            foreach (var target in FrequentTargetStore.LoadAll())
            {
                frequentByKey[target.Key] = target;
            }

            Text = "Quick Move";
            StartPosition = FormStartPosition.CenterScreen;
            MinimumSize = new Size(560, 420);
            Size = new Size(720, 520);
            Font = SystemFonts.MessageBoxFont;
            QuickMoveIcon.ApplyTo(this);

            var layout = new TableLayoutPanel
            {
                ColumnCount = 1,
                Dock = DockStyle.Fill,
                Padding = new Padding(12),
                RowCount = 4
            };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            textSearch = new TextBox
            {
                Dock = DockStyle.Fill,
                Margin = new Padding(0, 0, 0, 8)
            };
            textSearch.TextChanged += delegate { ApplyFilter(); };
            textSearch.KeyDown += HandleSearchKeyDown;

            listFolders = new ListBox
            {
                DisplayMember = nameof(FolderCandidate.DisplayPath),
                Dock = DockStyle.Fill,
                IntegralHeight = false,
                Margin = new Padding(0, 0, 0, 8)
            };
            listFolders.SelectedIndexChanged += delegate { UpdateOkState(); };
            listFolders.DoubleClick += delegate { ConfirmSelection(); };
            listFolders.KeyDown += HandleListKeyDown;

            checkMarkAsRead = new CheckBox
            {
                AutoSize = true,
                Checked = MarkAsReadSetting.Load(),
                Dock = DockStyle.Left,
                Margin = new Padding(0, 0, 0, 8),
                Text = "Mark as read before moving"
            };

            buttonOk = new Button
            {
                DialogResult = DialogResult.None,
                Enabled = false,
                Margin = new Padding(8, 0, 0, 0),
                Size = new Size(88, 28),
                Text = "OK"
            };
            buttonOk.Click += delegate { ConfirmSelection(); };

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
                Margin = new Padding(0),
                Padding = new Padding(0)
            };
            buttonPanel.Controls.Add(buttonCancel);
            buttonPanel.Controls.Add(buttonOk);

            layout.Controls.Add(textSearch, 0, 0);
            layout.Controls.Add(listFolders, 0, 1);
            layout.Controls.Add(checkMarkAsRead, 0, 2);
            layout.Controls.Add(buttonPanel, 0, 3);

            Controls.Add(layout);

            AcceptButton = buttonOk;
            CancelButton = buttonCancel;

            Shown += delegate { textSearch.Focus(); };
            ApplyFilter();
        }

        public string SelectedFolderEntryId { get; private set; }

        public string SelectedFolderStoreId { get; private set; }

        public bool MarkAsReadBeforeMoving
        {
            get { return checkMarkAsRead.Checked; }
        }

        private void ApplyFilter()
        {
            var terms = textSearch.Text
                .Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(term => term.Trim())
                .Where(term => term.Length > 0)
                .ToArray();

            // Frequent targets float to the top, ordered by usage (then recency); everything
            // else keeps its alphabetical order below them.
            var filtered = allFolders
                .Where(folder => MatchesAllTerms(folder.DisplayPath, terms))
                .OrderByDescending(FrequencyCount)
                .ThenByDescending(FrequencyTicks)
                .ThenBy(folder => folder.DisplayPath, StringComparer.OrdinalIgnoreCase)
                .ToList();

            listFolders.BeginUpdate();
            try
            {
                listFolders.DataSource = null;
                listFolders.DataSource = filtered;
                listFolders.DisplayMember = nameof(FolderCandidate.DisplayPath);

                if (filtered.Count > 0)
                {
                    listFolders.SelectedIndex = 0;
                }
            }
            finally
            {
                listFolders.EndUpdate();
            }

            UpdateOkState();
        }

        private FrequentTarget GetFrequency(FolderCandidate folder)
        {
            FrequentTarget target;
            return frequentByKey.TryGetValue(FrequentTarget.BuildKey(folder.StoreId, folder.EntryId), out target)
                ? target
                : null;
        }

        private int FrequencyCount(FolderCandidate folder)
        {
            var target = GetFrequency(folder);
            return target == null ? 0 : target.Count;
        }

        private long FrequencyTicks(FolderCandidate folder)
        {
            var target = GetFrequency(folder);
            return target == null ? 0L : target.LastUsedUtc.Ticks;
        }

        private static bool MatchesAllTerms(string displayPath, IEnumerable<string> terms)
        {
            foreach (var term in terms)
            {
                if (displayPath.IndexOf(term, StringComparison.OrdinalIgnoreCase) < 0)
                {
                    return false;
                }
            }

            return true;
        }

        private void UpdateOkState()
        {
            buttonOk.Enabled = listFolders.SelectedItem is FolderCandidate;
        }

        private void HandleSearchKeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Down)
            {
                MoveListSelection(1);
                e.Handled = true;
                e.SuppressKeyPress = true;
            }
            else if (e.KeyCode == Keys.Up)
            {
                MoveListSelection(-1);
                e.Handled = true;
                e.SuppressKeyPress = true;
            }
        }

        private void HandleListKeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                ConfirmSelection();
                e.Handled = true;
                e.SuppressKeyPress = true;
            }
        }

        private void MoveListSelection(int offset)
        {
            if (listFolders.Items.Count == 0)
            {
                return;
            }

            if (listFolders.SelectedIndex < 0)
            {
                listFolders.SelectedIndex = offset < 0 ? listFolders.Items.Count - 1 : 0;
                return;
            }

            var nextIndex = Math.Max(0, Math.Min(listFolders.Items.Count - 1, listFolders.SelectedIndex + offset));
            listFolders.SelectedIndex = nextIndex;
        }

        private void ConfirmSelection()
        {
            var candidate = listFolders.SelectedItem as FolderCandidate;
            if (candidate == null)
            {
                return;
            }

            // Remember the checkbox state so it is restored next time, even after an
            // Outlook restart. Persisted on confirm so it reflects the last move the user made.
            MarkAsReadSetting.Save(checkMarkAsRead.Checked);

            // Record this selection so the folder rises in the frequent-targets ordering.
            FrequentTargetStore.RecordUse(candidate.StoreId, candidate.EntryId, candidate.DisplayPath);

            SelectedFolderEntryId = candidate.EntryId;
            SelectedFolderStoreId = candidate.StoreId;
            DialogResult = DialogResult.OK;
            Close();
        }
    }
}
