using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;
using ExcelCSIToolBox.Application.Interfaces.Etabs;
using ExcelCSIToolBox.Core.Common.Results;
using ExcelCSIToolBox.Core.Contracts.CSI;

namespace ExcelCSIToolBoxAddIn.UI.Forms
{
    internal partial class ShellUniformLoadSetSelectionForm : Form
    {
        private readonly IEtabsShellUniformLoadSetSelectionService _selectionService;
        private readonly HashSet<string> _checkedLoadSetNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private readonly HashSet<string> _checkedStoryNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private List<string> _allLoadSetNames = new List<string>();
        private List<string> _allStoryNames = new List<string>();
        private bool _hasLoadedInitialLoadSets;
        private bool _hasPopulatedStoryNames;
        private bool _isBusy;
        private bool _allowCloseWhileBusy;
        private bool _updatingCheckState;
        private bool _updatingStoryCheckState;

        public ShellUniformLoadSetSelectionForm(IEtabsShellUniformLoadSetSelectionService selectionService)
        {
            _selectionService = selectionService ?? throw new ArgumentNullException(nameof(selectionService));
            InitializeComponent();
        }

        private void ShellUniformLoadSetSelectionForm_Load(object sender, EventArgs e)
        {
            lblStatus.Text = "Loading Shell Uniform Load Sets and Stories...";
            UpdateSelectionStatus();
        }

        private void ShellUniformLoadSetSelectionForm_Shown(object sender, EventArgs e)
        {
            if (_hasLoadedInitialLoadSets)
            {
                return;
            }

            _hasLoadedInitialLoadSets = true;
            BeginInvoke(new MethodInvoker(RefreshLoadSetNames));
        }

        private void txtFilter_TextChanged(object sender, EventArgs e)
        {
            ApplyFilter();
        }

        private void chkLoadSets_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            if (_updatingCheckState || e.Index < 0 || e.Index >= chkLoadSets.Items.Count)
            {
                return;
            }

            string loadSetName = Convert.ToString(chkLoadSets.Items[e.Index], CultureInfo.CurrentCulture);
            if (string.IsNullOrWhiteSpace(loadSetName))
            {
                return;
            }

            if (e.NewValue == CheckState.Checked)
            {
                _checkedLoadSetNames.Add(loadSetName.Trim());
            }
            else
            {
                _checkedLoadSetNames.Remove(loadSetName.Trim());
            }

            BeginInvoke(new MethodInvoker(UpdateSelectionStatus));
        }

        private void btnSelectAll_Click(object sender, EventArgs e)
        {
            _updatingCheckState = true;
            try
            {
                for (int i = 0; i < chkLoadSets.Items.Count; i++)
                {
                    string loadSetName = Convert.ToString(chkLoadSets.Items[i], CultureInfo.CurrentCulture);
                    if (!string.IsNullOrWhiteSpace(loadSetName))
                    {
                        _checkedLoadSetNames.Add(loadSetName.Trim());
                    }

                    chkLoadSets.SetItemChecked(i, true);
                }
            }
            finally
            {
                _updatingCheckState = false;
            }

            UpdateSelectionStatus();
        }

        private void btnClearSelection_Click(object sender, EventArgs e)
        {
            _checkedLoadSetNames.Clear();
            _updatingCheckState = true;
            try
            {
                for (int i = 0; i < chkLoadSets.Items.Count; i++)
                {
                    chkLoadSets.SetItemChecked(i, false);
                }
            }
            finally
            {
                _updatingCheckState = false;
            }

            UpdateSelectionStatus();
        }

        private void btnRefresh_Click(object sender, EventArgs e)
        {
            RefreshLoadSetNames();
        }

        private void btnSelectShells_Click(object sender, EventArgs e)
        {
            List<string> selectedLoadSets = _checkedLoadSetNames
                .OrderBy(name => name, StringComparer.OrdinalIgnoreCase)
                .ToList();
            List<string> selectedStoryNames = GetSelectedStoryNames();
            if (selectedLoadSets.Count == 0)
            {
                MessageBox.Show(this, "Select at least one Shell Uniform Load Set.", "Select Shells", MessageBoxButtons.OK, MessageBoxIcon.Information);
                UpdateSelectionStatus();
                return;
            }

            if (selectedStoryNames.Count == 0)
            {
                MessageBox.Show(this, "Select at least one ETABS Story.", "Select Shells", MessageBoxButtons.OK, MessageBoxIcon.Information);
                UpdateSelectionStatus();
                return;
            }

            btnSelectShells.Enabled = false;
            Cursor previousCursor = Cursor.Current;
            Cursor.Current = Cursors.WaitCursor;
            SetBusy(true, "Selecting ETABS shell objects...");
            ResetSelectionProgress(true, "Preparing selection...");
            try
            {
                OperationResult<ShellUniformLoadSetSelectionResultDto> result =
                    _selectionService.SelectShellsByLoadSets(
                        selectedLoadSets,
                        selectedStoryNames,
                        new ImmediateProgress<ShellUniformLoadSetSelectionProgressDto>(UpdateSelectionProgress));
                if (!result.IsSuccess)
                {
                    lblStatus.Text = result.Message;
                    MessageBox.Show(this, result.Message, "Select Shells", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                ShellUniformLoadSetSelectionResultDto data = result.Data ?? new ShellUniformLoadSetSelectionResultDto { Message = result.Message };
                string message = string.IsNullOrWhiteSpace(data.Message) ? result.Message : data.Message;
                if (!string.IsNullOrWhiteSpace(data.WarningMessage))
                {
                    message += "\r\n\r\n" + data.WarningMessage;
                }

                MessageBox.Show(this, message, "Select Shells", MessageBoxButtons.OK, MessageBoxIcon.Information);
                _allowCloseWhileBusy = true;
                DialogResult = DialogResult.OK;
                Close();
            }
            finally
            {
                Cursor.Current = previousCursor;
                if (!IsDisposed)
                {
                    SetBusy(false, lblStatus.Text);
                    if (DialogResult != DialogResult.OK)
                    {
                        ResetSelectionProgress(false, lblStatus.Text);
                    }

                    UpdateSelectionStatus();
                }
            }
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (_isBusy && !_allowCloseWhileBusy)
            {
                e.Cancel = true;
                return;
            }

            base.OnFormClosing(e);
        }

        private void RefreshLoadSetNames()
        {
            Cursor previousCursor = Cursor.Current;
            Cursor.Current = Cursors.WaitCursor;
            SetBusy(true, "Loading Shell Uniform Load Sets and Stories...");
            try
            {
                OperationResult<IReadOnlyList<string>> result = _selectionService.GetLoadSetNames();
                if (!result.IsSuccess)
                {
                    _allLoadSetNames = new List<string>();
                    _allStoryNames = new List<string>();
                    _checkedStoryNames.Clear();
                    _hasPopulatedStoryNames = false;
                    PopulateStoryList();
                    ApplyFilter();
                    lblStatus.Text = result.Message;
                    MessageBox.Show(this, result.Message, "Shell Uniform Load Sets", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                _allLoadSetNames = (result.Data ?? new string[0])
                    .Where(name => !string.IsNullOrWhiteSpace(name))
                    .Select(name => name.Trim())
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .OrderBy(name => name, StringComparer.OrdinalIgnoreCase)
                    .ToList();

                OperationResult<IReadOnlyList<string>> storyResult = _selectionService.GetStoryNames();
                if (storyResult.IsSuccess)
                {
                    _allStoryNames = NormalizeStoryNames(storyResult.Data);
                    PopulateStoryList();
                }
                else
                {
                    _allStoryNames = new List<string>();
                    _checkedStoryNames.Clear();
                    _hasPopulatedStoryNames = false;
                    PopulateStoryList();
                    MessageBox.Show(this, storyResult.Message, "ETABS Stories", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }

                HashSet<string> validNames = new HashSet<string>(_allLoadSetNames, StringComparer.OrdinalIgnoreCase);
                _checkedLoadSetNames.RemoveWhere(name => !validNames.Contains(name));
                ApplyFilter();
                lblStatus.Text = storyResult.IsSuccess
                    ? result.Message + " " + storyResult.Message
                    : result.Message;
                UpdateSelectionStatus();
            }
            finally
            {
                Cursor.Current = previousCursor;
                SetBusy(false, lblStatus.Text);
            }
        }

        private void ApplyFilter()
        {
            string filter = (txtFilter.Text ?? string.Empty).Trim();
            List<string> filtered = _allLoadSetNames
                .Where(name => string.IsNullOrWhiteSpace(filter) || name.IndexOf(filter, StringComparison.OrdinalIgnoreCase) >= 0)
                .ToList();

            _updatingCheckState = true;
            try
            {
                chkLoadSets.Items.Clear();
                foreach (string loadSetName in filtered)
                {
                    int index = chkLoadSets.Items.Add(loadSetName);
                    chkLoadSets.SetItemChecked(index, _checkedLoadSetNames.Contains(loadSetName));
                }
            }
            finally
            {
                _updatingCheckState = false;
            }

            UpdateEmptyState(filtered.Count);
            UpdateSelectionStatus();
        }

        private void chkStories_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            if (_updatingStoryCheckState || e.Index < 0 || e.Index >= chkStories.Items.Count)
            {
                return;
            }

            string storyName = Convert.ToString(chkStories.Items[e.Index], CultureInfo.CurrentCulture);
            if (string.IsNullOrWhiteSpace(storyName))
            {
                return;
            }

            if (e.NewValue == CheckState.Checked)
            {
                _checkedStoryNames.Add(storyName.Trim());
            }
            else
            {
                _checkedStoryNames.Remove(storyName.Trim());
            }

            BeginInvoke(new MethodInvoker(UpdateSelectionStatus));
        }

        private void btnSelectAllStories_Click(object sender, EventArgs e)
        {
            _updatingStoryCheckState = true;
            try
            {
                for (int i = 0; i < chkStories.Items.Count; i++)
                {
                    string storyName = Convert.ToString(chkStories.Items[i], CultureInfo.CurrentCulture);
                    if (!string.IsNullOrWhiteSpace(storyName))
                    {
                        _checkedStoryNames.Add(storyName.Trim());
                    }

                    chkStories.SetItemChecked(i, true);
                }
            }
            finally
            {
                _updatingStoryCheckState = false;
            }

            UpdateSelectionStatus();
        }

        private void btnClearStories_Click(object sender, EventArgs e)
        {
            _checkedStoryNames.Clear();
            _updatingStoryCheckState = true;
            try
            {
                for (int i = 0; i < chkStories.Items.Count; i++)
                {
                    chkStories.SetItemChecked(i, false);
                }
            }
            finally
            {
                _updatingStoryCheckState = false;
            }

            UpdateSelectionStatus();
        }

        private void UpdateEmptyState(int filteredCount)
        {
            if (_allLoadSetNames.Count == 0)
            {
                lblEmptyState.Text = "No Shell Uniform Load Sets exist in the connected ETABS model.";
                lblEmptyState.Visible = true;
                return;
            }

            if (filteredCount == 0)
            {
                lblEmptyState.Text = "No load sets match the current filter.";
                lblEmptyState.Visible = true;
                return;
            }

            lblEmptyState.Visible = false;
        }

        private void UpdateSelectionStatus()
        {
            int selectedCount = _checkedLoadSetNames.Count;
            lblSelectionStatus.Text = selectedCount.ToString(CultureInfo.InvariantCulture) + " load set(s) selected";
            int selectedStoryCount = _checkedStoryNames.Count;
            lblStorySelectionStatus.Text = selectedStoryCount.ToString(CultureInfo.InvariantCulture) + " story(s) selected";
            if (_isBusy)
            {
                btnSelectShells.Enabled = false;
                btnSelectAll.Enabled = false;
                btnClearSelection.Enabled = false;
                btnSelectAllStories.Enabled = false;
                btnClearStories.Enabled = false;
                btnRefresh.Enabled = false;
                btnCancel.Enabled = false;
                txtFilter.Enabled = false;
                chkStories.Enabled = false;
                chkLoadSets.Enabled = false;
                return;
            }

            txtFilter.Enabled = true;
            chkStories.Enabled = true;
            chkLoadSets.Enabled = true;
            btnRefresh.Enabled = true;
            btnSelectShells.Enabled = selectedCount > 0 && selectedStoryCount > 0;
            btnSelectAll.Enabled = chkLoadSets.Items.Count > 0;
            btnClearSelection.Enabled = selectedCount > 0;
            btnSelectAllStories.Enabled = chkStories.Items.Count > 0;
            btnClearStories.Enabled = selectedStoryCount > 0;
            btnCancel.Enabled = true;
        }

        private void SetBusy(bool isBusy, string status)
        {
            _isBusy = isBusy;
            if (!string.IsNullOrWhiteSpace(status))
            {
                lblStatus.Text = status;
            }

            UpdateSelectionStatus();
        }

        private void ResetSelectionProgress(bool visible, string status)
        {
            progressSelection.Style = ProgressBarStyle.Continuous;
            progressSelection.Minimum = 0;
            progressSelection.Maximum = 100;
            progressSelection.Value = 0;
            progressSelection.Visible = visible;
            if (!string.IsNullOrWhiteSpace(status))
            {
                lblStatus.Text = status;
            }

            progressSelection.Refresh();
            lblStatus.Refresh();
        }

        private void UpdateSelectionProgress(ShellUniformLoadSetSelectionProgressDto progress)
        {
            if (progressSelection.IsDisposed || IsDisposed)
            {
                return;
            }

            if (InvokeRequired)
            {
                BeginInvoke(new Action<ShellUniformLoadSetSelectionProgressDto>(UpdateSelectionProgress), progress);
                return;
            }

            if (progress == null)
            {
                return;
            }

            progressSelection.Visible = true;
            if (!string.IsNullOrWhiteSpace(progress.Message))
            {
                lblStatus.Text = progress.Message;
            }

            if (progress.IsIndeterminate || progress.Total <= 0)
            {
                progressSelection.Style = ProgressBarStyle.Marquee;
            }
            else
            {
                progressSelection.Style = ProgressBarStyle.Continuous;
                progressSelection.Minimum = 0;
                progressSelection.Maximum = Math.Max(1, progress.Total);
                progressSelection.Value = Math.Max(0, Math.Min(progress.Current, progressSelection.Maximum));
            }

            progressSelection.Refresh();
            lblStatus.Refresh();
            Application.DoEvents();
        }

        private void PopulateStoryList()
        {
            HashSet<string> previousStoryNames = new HashSet<string>(_checkedStoryNames, StringComparer.OrdinalIgnoreCase);
            bool selectAllStories = !_hasPopulatedStoryNames;
            chkStories.BeginUpdate();
            _updatingStoryCheckState = true;
            try
            {
                chkStories.Items.Clear();
                _checkedStoryNames.Clear();
                foreach (string storyName in _allStoryNames)
                {
                    int index = chkStories.Items.Add(storyName);
                    bool isChecked = selectAllStories || previousStoryNames.Contains(storyName);
                    if (isChecked)
                    {
                        _checkedStoryNames.Add(storyName);
                    }

                    chkStories.SetItemChecked(index, isChecked);
                }
            }
            finally
            {
                _updatingStoryCheckState = false;
                chkStories.EndUpdate();
            }

            if (_allStoryNames.Count > 0)
            {
                _hasPopulatedStoryNames = true;
            }

            UpdateSelectionStatus();
        }

        private List<string> GetSelectedStoryNames()
        {
            return _checkedStoryNames
                .OrderBy(name => name, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private static List<string> NormalizeStoryNames(IEnumerable<string> storyNames)
        {
            HashSet<string> seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            List<string> names = new List<string>();
            foreach (string rawName in storyNames ?? new string[0])
            {
                string name = string.IsNullOrWhiteSpace(rawName) ? string.Empty : rawName.Trim();
                if (string.IsNullOrWhiteSpace(name) || !seen.Add(name))
                {
                    continue;
                }

                names.Add(name);
            }

            return names;
        }

        private sealed class ImmediateProgress<T> : IProgress<T>
        {
            private readonly Action<T> _handler;

            public ImmediateProgress(Action<T> handler)
            {
                _handler = handler ?? throw new ArgumentNullException(nameof(handler));
            }

            public void Report(T value)
            {
                _handler(value);
            }
        }
    }
}
