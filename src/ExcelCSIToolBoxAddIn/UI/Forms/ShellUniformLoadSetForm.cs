using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;
using ExcelCSIToolBox.Core.Abstractions.CSI;
using ExcelCSIToolBox.Core.Common.Results;
using ExcelCSIToolBox.Core.Contracts.CSI;

namespace ExcelCSIToolBoxAddIn.UI.Forms
{
    internal partial class ShellUniformLoadSetForm : Form
    {
        private const string NameColumnKey = "UniformLoadSetName";
        private const string SuggestedNameColumnKey = "SuggestedName";
        private const int NameColumnMinimumWidth = 180;
        private const int NameColumnMaximumWidth = 320;
        private const int LoadPatternColumnMinimumWidth = 70;
        private const int LoadPatternColumnMaximumWidth = 160;
        private readonly ICSISapModelConnectionService _connectionService;
        private readonly ICsiApiDispatcher _csiApiDispatcher;
        private readonly Action<IntPtr> _exportCurrentDefinitionsAction;
        private readonly ExcelSelectedRangeReader _excelRangeReader;
        private readonly Dictionary<string, string> _loadPatternLookup = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        private readonly List<UnitOption> _unitOptions = new List<UnitOption>();

        public ShellUniformLoadSetForm(
            ICSISapModelConnectionService connectionService,
            Action<IntPtr> exportCurrentDefinitionsAction = null,
            ICsiApiDispatcher csiApiDispatcher = null)
        {
            if (connectionService == null) throw new ArgumentNullException(nameof(connectionService));
            _connectionService = connectionService;
            _csiApiDispatcher = csiApiDispatcher ?? new ExcelCSIToolBox.Infrastructure.CSISapModel.CurrentThreadCsiApiDispatcher();
            _exportCurrentDefinitionsAction = exportCurrentDefinitionsAction;
            _excelRangeReader = new ExcelSelectedRangeReader();
            InitializeComponent();
            InitializeUnitOptions();
            InitializeGrid();
        }

        private void ShellUniformLoadSetForm_Load(object sender, EventArgs e)
        {
            LoadEtabsContext();
        }

        private void LoadEtabsContext()
        {
            OperationResult<ShellUniformLoadSetContextDto> result = _connectionService.GetShellUniformLoadSetContext();
            if (!result.IsSuccess)
            {
                SetStatus("Status: ETABS connection is not ready.");
                MessageBox.Show(this, result.Message, "ETABS Connection Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                btnApply.Enabled = false;
                btnImportExcelRange.Enabled = false;
                btnAddLoadPattern.Enabled = false;
                btnExportDefinitions.Enabled = false;
                return;
            }

            btnApply.Enabled = true;
            btnImportExcelRange.Enabled = true;
            btnExportDefinitions.Enabled = _exportCurrentDefinitionsAction != null;

            ShellUniformLoadSetContextDto context = result.Data ?? new ShellUniformLoadSetContextDto();
            lblModelName.Text = "ETABS Model: " + (string.IsNullOrWhiteSpace(context.ModelFileName) ? "-" : context.ModelFileName);
            lblPresentUnits.Text = "Present Units: " + (string.IsNullOrWhiteSpace(context.PresentUnitsText) ? "-" : context.PresentUnitsText);

            _loadPatternLookup.Clear();
            foreach (string name in context.LoadPatternNames ?? new List<string>())
            {
                string trimmed = NormalizeText(name);
                if (!string.IsNullOrWhiteSpace(trimmed) && !_loadPatternLookup.ContainsKey(trimmed))
                {
                    _loadPatternLookup.Add(trimmed, trimmed);
                }
            }

            SelectCurrentOrDefaultUnit();
            LoadCurrentModelDefinitions();
            UpdateLoadPatternButtonState();
        }

        private bool LoadCurrentModelDefinitions()
        {
            OperationResult<IReadOnlyList<ShellUniformLoadSetDefinitionDto>> definitionsResult = _connectionService.GetShellUniformLoadSetDefinitions();
            if (!definitionsResult.IsSuccess)
            {
                SetStatus("Status: Could not load current Shell Uniform Load Sets.");
                MessageBox.Show(this, definitionsResult.Message, "Shell Uniform Load Set Read Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            IReadOnlyList<ShellUniformLoadSetDefinitionDto> definitions = definitionsResult.Data ?? new List<ShellUniformLoadSetDefinitionDto>();
            PopulateDefinitions(definitions);

            if (definitions.Count == 0)
            {
                SetStatus("Status: No existing Shell Uniform Load Sets found.");
                return true;
            }

            SetStatus("Status: Loaded " + definitions.Count.ToString(CultureInfo.InvariantCulture) + " Shell Uniform Load Set(s) from ETABS.");
            return true;
        }

        private void btnRefreshDefinitions_Click(object sender, EventArgs e)
        {
            if (GridContainsData())
            {
                DialogResult confirm = MessageBox.Show(
                    this,
                    "Refresh current ETABS definitions?\r\n\r\nThis will replace the table shown in the manager. Any unsaved edits in the grid will be lost.",
                    "Refresh Current Definitions",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);
                if (confirm != DialogResult.Yes)
                {
                    return;
                }
            }

            btnRefreshDefinitions.Enabled = false;
            Cursor previousCursor = Cursor.Current;
            Cursor.Current = Cursors.WaitCursor;
            try
            {
                SetStatus("Status: Refreshing current ETABS definitions...");
                LoadEtabsContext();
            }
            finally
            {
                Cursor.Current = previousCursor;
                btnRefreshDefinitions.Enabled = true;
            }
        }

        private void btnExportDefinitions_Click(object sender, EventArgs e)
        {
            if (_exportCurrentDefinitionsAction == null)
            {
                MessageBox.Show(this, "Excel export is not available from this manager instance.", "Export Current Definitions", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            try
            {
                _exportCurrentDefinitionsAction(Handle);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Failed to export current Shell Uniform Load Set definitions:\r\n\r\n" + ex.Message, "Export Current Definitions", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void PopulateDefinitions(IReadOnlyList<ShellUniformLoadSetDefinitionDto> definitions)
        {
            InitializeGrid();
            if (definitions == null || definitions.Count == 0)
            {
                return;
            }

            List<string> patternNames = definitions
                .Where(definition => definition != null && definition.LoadValuesByPattern != null)
                .SelectMany(definition => definition.LoadValuesByPattern.Keys)
                .Where(name => !string.IsNullOrWhiteSpace(name))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(name => name, StringComparer.OrdinalIgnoreCase)
                .ToList();

            foreach (string patternName in patternNames)
            {
                AddLoadPatternColumn(patternName);
            }

            foreach (ShellUniformLoadSetDefinitionDto definition in definitions.Where(definition => definition != null))
            {
                int rowIndex = dgvLoadSets.Rows.Add();
                DataGridViewRow row = dgvLoadSets.Rows[rowIndex];
                row.Cells[NameColumnKey].Value = definition.Name ?? string.Empty;

                if (definition.LoadValuesByPattern == null)
                {
                    continue;
                }

                foreach (KeyValuePair<string, double> loadValue in definition.LoadValuesByPattern)
                {
                    DataGridViewColumn column = FindLoadPatternColumn(loadValue.Key);
                    if (column == null)
                    {
                        AddLoadPatternColumn(loadValue.Key);
                        column = FindLoadPatternColumn(loadValue.Key);
                    }

                    if (column != null)
                    {
                        row.Cells[column.Index].Value = FormatGridValue(loadValue.Value);
                    }
                }
            }

            ApplySuggestedNamesIfEnabled();
            AutoFitGrid();
        }

        private void InitializeUnitOptions()
        {
            _unitOptions.Clear();
            _unitOptions.Add(new UnitOption("N-mm", 3, 4, 2));
            _unitOptions.Add(new UnitOption("kN-m", 4, 6, 2));
            _unitOptions.Add(new UnitOption("kip-ft", 2, 2, 2));
            _unitOptions.Add(new UnitOption("lb-in", 1, 1, 2));

            cboSelectedUnits.Items.Clear();
            foreach (UnitOption option in _unitOptions)
            {
                cboSelectedUnits.Items.Add(option);
            }

            if (cboSelectedUnits.Items.Count > 0)
            {
                cboSelectedUnits.SelectedIndex = 0;
            }
        }

        private void SelectCurrentOrDefaultUnit()
        {
            OperationResult<CSISapModelPresentUnitSystemDTO> unitResult = _connectionService.GetPresentUnitSystem();
            if (unitResult.IsSuccess && unitResult.Data != null)
            {
                for (int i = 0; i < _unitOptions.Count; i++)
                {
                    if (_unitOptions[i].Matches(unitResult.Data))
                    {
                        cboSelectedUnits.SelectedIndex = i;
                        return;
                    }
                }
            }

            UnitOption defaultOption = _unitOptions.FirstOrDefault(option => string.Equals(option.DisplayName, "N-mm", StringComparison.OrdinalIgnoreCase));
            if (defaultOption != null)
            {
                cboSelectedUnits.SelectedItem = defaultOption;
            }
        }

        private bool IsLoadPatternColumn(int columnIndex)
        {
            if (columnIndex < 0 || columnIndex >= dgvLoadSets.Columns.Count) return false;
            var col = dgvLoadSets.Columns[columnIndex];
            return col.Name != NameColumnKey && col.Name != SuggestedNameColumnKey;
        }

        private bool IsLoadPatternColumn(DataGridViewColumn column)
        {
            return column != null && column.Name != NameColumnKey && column.Name != SuggestedNameColumnKey;
        }

        private void InitializeGrid()
        {
            ConfigureGridTextWrapping();
            dgvLoadSets.Columns.Clear();
            dgvLoadSets.Rows.Clear();
            
            // Name Column (always editable)
            DataGridViewTextBoxColumn nameColumn = new DataGridViewTextBoxColumn
            {
                Name = NameColumnKey,
                HeaderText = NameColumnKey,
                Tag = NameColumnKey,
                SortMode = DataGridViewColumnSortMode.NotSortable,
                MinimumWidth = NameColumnMinimumWidth,
                Width = 220,
                Frozen = true,
                ReadOnly = false
            };
            nameColumn.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            nameColumn.HeaderCell.Style.WrapMode = DataGridViewTriState.True;
            dgvLoadSets.Columns.Add(nameColumn);

            // Suggested Name Column (read-only, visibility toggled by checkbox)
            DataGridViewTextBoxColumn suggestedNameColumn = new DataGridViewTextBoxColumn
            {
                Name = SuggestedNameColumnKey,
                HeaderText = "Suggested Name",
                Tag = SuggestedNameColumnKey,
                SortMode = DataGridViewColumnSortMode.NotSortable,
                MinimumWidth = NameColumnMinimumWidth,
                Width = 220,
                Frozen = true,
                ReadOnly = true,
                Visible = chkApplySuggestedName != null && chkApplySuggestedName.Checked
            };
            suggestedNameColumn.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            suggestedNameColumn.HeaderCell.Style.WrapMode = DataGridViewTriState.True;
            dgvLoadSets.Columns.Add(suggestedNameColumn);

            AutoFitGrid();
        }

        private void btnImportExcelRange_Click(object sender, EventArgs e)
        {
            OperationResult<ExcelSelectedRangeData> readResult;
            
            // Bring Excel to the foreground before hiding the form to make the transition seamless
            _excelRangeReader.ActivateExcel();
            Application.DoEvents();
            System.Threading.Thread.Sleep(50);

            Hide();
            Application.DoEvents();

            try
            {
                readResult = _excelRangeReader.ReadSelectedRange();
            }
            finally
            {
                Show();
                Activate();
            }

            if (!readResult.IsSuccess)
            {
                MessageBox.Show(this, readResult.Message, "Excel Selection Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            ExcelSelectedRangeData range = readResult.Data;
            if (range.ColumnCount < 2 || range.RowCount < 2)
            {
                MessageBox.Show(this, "A valid import must contain a header row, at least one data row, and at least two columns.", "Excel Selection Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            OperationResult<List<string>> headerResult = ValidateImportHeaders(range);
            if (!headerResult.IsSuccess)
            {
                MessageBox.Show(this, headerResult.Message, "Header Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            TransferExcelDataToGrid(range, headerResult.Data);
            SetStatus("Status: " + CountNonBlankRows().ToString(CultureInfo.InvariantCulture) + " load sets ready.");
        }

        private OperationResult<List<string>> ValidateImportHeaders(ExcelSelectedRangeData range)
        {
            var canonicalHeaders = new List<string>();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var duplicateHeaders = new List<string>();
            var invalidHeaders = new List<string>();

            for (int column = 2; column <= range.ColumnCount; column++)
            {
                string header = NormalizeText(ToCellText(range.GetValue(1, column)));
                if (string.IsNullOrWhiteSpace(header))
                {
                    invalidHeaders.Add("(blank column " + column.ToString(CultureInfo.InvariantCulture) + ")");
                    continue;
                }

                if (!seen.Add(header))
                {
                    duplicateHeaders.Add(header);
                    continue;
                }

                string canonical;
                if (!_loadPatternLookup.TryGetValue(header, out canonical))
                {
                    invalidHeaders.Add(header);
                    continue;
                }

                canonicalHeaders.Add(canonical);
            }

            if (duplicateHeaders.Count > 0)
            {
                return OperationResult<List<string>>.Failure(
                    "Duplicate Load Pattern headers were found:\r\n\r\n" +
                    string.Join("\r\n", duplicateHeaders.Select(header => "- " + header)) +
                    "\r\n\r\nEach Load Pattern may appear only once.");
            }

            if (invalidHeaders.Count > 0)
            {
                return OperationResult<List<string>>.Failure(
                    "Cannot import the selected Excel range.\r\n\r\n" +
                    "The following headers are not valid Load Patterns in the current ETABS model:\r\n\r\n" +
                    string.Join("\r\n", invalidHeaders.Select(header => "- " + header)) +
                    "\r\n\r\nCorrect the Excel headers and try again.");
            }

            return canonicalHeaders.Count == 0
                ? OperationResult<List<string>>.Failure("Cannot import the selected Excel range.\r\n\r\nAt least one Load Pattern column is required.")
                : OperationResult<List<string>>.Success(canonicalHeaders);
        }

        private void TransferExcelDataToGrid(ExcelSelectedRangeData range, IReadOnlyList<string> canonicalHeaders)
        {
            ConfigureGridTextWrapping();
            foreach (string header in canonicalHeaders)
            {
                AddLoadPatternColumn(header);
            }

            for (int row = 2; row <= range.RowCount; row++)
            {
                if (IsBlankExcelRow(range, row))
                {
                    continue;
                }

                int gridRowIndex = dgvLoadSets.Rows.Add();
                DataGridViewRow gridRow = dgvLoadSets.Rows[gridRowIndex];
                
                // First column in Excel is the load set name
                gridRow.Cells[NameColumnKey].Value = ToCellText(range.GetValue(row, 1));
                
                // Subsequent columns are load pattern values
                for (int column = 2; column <= range.ColumnCount; column++)
                {
                    string header = canonicalHeaders[column - 2];
                    gridRow.Cells[header].Value = ToCellText(range.GetValue(row, column));
                }
            }

            ApplySuggestedNamesIfEnabled();
            AutoFitGrid();
            UpdateLoadPatternButtonState();
        }

        private void btnAddRow_Click(object sender, EventArgs e)
        {
            dgvLoadSets.Rows.Add();
            ApplySuggestedNamesIfEnabled();
            AutoFitGrid();
            SetStatus("Status: " + CountNonBlankRows().ToString(CultureInfo.InvariantCulture) + " load sets ready.");
        }

        private void btnDeleteRow_Click(object sender, EventArgs e)
        {
            var rows = dgvLoadSets.SelectedCells
                .Cast<DataGridViewCell>()
                .Select(cell => cell.OwningRow)
                .Where(row => row != null && !row.IsNewRow)
                .Distinct()
                .ToList();

            if (rows.Count == 0 && dgvLoadSets.CurrentRow != null && !dgvLoadSets.CurrentRow.IsNewRow)
            {
                rows.Add(dgvLoadSets.CurrentRow);
            }

            if (rows.Count == 0)
            {
                return;
            }

            if (rows.Count > 1)
            {
                DialogResult confirm = MessageBox.Show(this, "Delete the selected rows?", "Delete Rows", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (confirm != DialogResult.Yes)
                {
                    return;
                }
            }

            foreach (DataGridViewRow row in rows)
            {
                dgvLoadSets.Rows.Remove(row);
            }

            ApplySuggestedNamesIfEnabled();
            AutoFitGrid();
            SetStatus("Status: " + CountNonBlankRows().ToString(CultureInfo.InvariantCulture) + " load sets ready.");
        }

        private void btnAddLoadPattern_Click(object sender, EventArgs e)
        {
            List<string> available = _loadPatternLookup.Values
                .Where(name => !GridHasLoadPatternColumn(name))
                .OrderBy(name => name, StringComparer.OrdinalIgnoreCase)
                .ToList();

            if (available.Count == 0)
            {
                MessageBox.Show(this, "All available ETABS Load Patterns have already been added.", "Add Load Pattern", MessageBoxButtons.OK, MessageBoxIcon.Information);
                UpdateLoadPatternButtonState();
                return;
            }

            using (var picker = new LoadPatternPickerForm(available))
            {
                if (picker.ShowDialog(this) != DialogResult.OK || string.IsNullOrWhiteSpace(picker.SelectedLoadPattern))
                {
                    return;
                }

                AddLoadPatternColumn(picker.SelectedLoadPattern);
                ApplySuggestedNamesIfEnabled();
                AutoFitGrid();
                UpdateLoadPatternButtonState();
            }
        }

        private void btnRemoveLoadPattern_Click(object sender, EventArgs e)
        {
            DataGridViewColumn column = GetSelectedLoadPatternColumn();
            if (column == null)
            {
                MessageBox.Show(this, "Select a Load Pattern column to remove.", "Remove Load Pattern", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            bool hasValues = dgvLoadSets.Rows.Cast<DataGridViewRow>()
                .Where(row => !row.IsNewRow)
                .Any(row => !string.IsNullOrWhiteSpace(ToCellText(row.Cells[column.Index].Value)));

            if (hasValues)
            {
                DialogResult confirm = MessageBox.Show(
                    this,
                    "Load Pattern \"" + column.HeaderText + "\" contains load values.\r\n\r\nRemoving this column will delete all values in the column.\r\n\r\nContinue?",
                    "Remove Load Pattern",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Warning);
                if (confirm != DialogResult.Yes)
                {
                    return;
                }
            }

            dgvLoadSets.Columns.Remove(column);
            ApplySuggestedNamesIfEnabled();
            AutoFitGrid();
            UpdateLoadPatternButtonState();
        }

        private void btnClear_Click(object sender, EventArgs e)
        {
            if (GridContainsData())
            {
                DialogResult confirm = MessageBox.Show(this, "Clear all rows and Load Pattern columns?", "Clear Table", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (confirm != DialogResult.Yes)
                {
                    return;
                }
            }

            InitializeGrid();
            AutoFitGrid();
            UpdateLoadPatternButtonState();
            SetStatus("Status: Ready.");
        }

        private void btnApply_Click(object sender, EventArgs e)
        {
            ApplySuggestedNamesIfEnabled();
            OperationResult<List<ShellUniformLoadSetDefinitionDto>> parseResult = ValidateAndCreateDefinitions();
            if (!parseResult.IsSuccess)
            {
                SetStatus("Status: Fix highlighted cells and try again.");
                MessageBox.Show(this, parseResult.Message, "Grid Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            UnitOption selectedUnitOption = cboSelectedUnits.SelectedItem as UnitOption;
            if (selectedUnitOption == null)
            {
                SetStatus("Status: Please select an ETABS unit system.");
                MessageBox.Show(this, "Please select an ETABS unit system.", "Unit Selection Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            var unitSystemDto = selectedUnitOption.ToDto();
            var definitionsToApply = parseResult.Data;

            btnApply.Enabled = false;
            Cursor previousCursor = Cursor.Current;
            Cursor.Current = Cursors.WaitCursor;

            OperationResult unitResult = null;
            OperationResult<ShellUniformLoadSetApplyResultDto> applyResult = null;
            OperationResult<IReadOnlyList<ShellUniformLoadSetDefinitionDto>> refreshResult = null;

            try
            {
                using (var progressForm = new ProgressForm("Applying Changes", "Applying changes to ETABS, please wait..."))
                {
                    progressForm.Show(this);
                    progressForm.Refresh();

                    _csiApiDispatcher.Invoke(() =>
                    {
                        try
                        {
                            unitResult = _connectionService.SetPresentUnitSystem(unitSystemDto);
                            if (unitResult.IsSuccess)
                            {
                                applyResult = _connectionService.ApplyShellUniformLoadSets(definitionsToApply);
                                if (applyResult.IsSuccess)
                                {
                                    refreshResult = _connectionService.GetShellUniformLoadSetDefinitions();
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            applyResult = OperationResult<ShellUniformLoadSetApplyResultDto>.Failure("An error occurred: " + ex.Message);
                        }
                        finally
                        {
                            progressForm.Close();
                        }
                    });
                }

                if (unitResult != null && !unitResult.IsSuccess)
                {
                    SetStatus("Status: Could not set ETABS units.");
                    MessageBox.Show(this, unitResult.Message, "Unit Selection Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                if (applyResult == null || !applyResult.IsSuccess)
                {
                    SetStatus("Status: ETABS table update failed.");
                    string errMsg = applyResult != null ? applyResult.Message : "ETABS application did not respond.";
                    MessageBox.Show(this, errMsg, "ETABS Import Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                if (refreshResult != null && refreshResult.IsSuccess)
                {
                    IReadOnlyList<ShellUniformLoadSetDefinitionDto> definitions = refreshResult.Data ?? new List<ShellUniformLoadSetDefinitionDto>();
                    PopulateDefinitions(definitions);
                    UpdateLoadPatternButtonState();

                    if (definitions.Count == 0)
                    {
                        SetStatus("Status: No existing Shell Uniform Load Sets found.");
                    }
                    else
                    {
                        SetStatus("Status: Loaded " + definitions.Count.ToString(CultureInfo.InvariantCulture) + " Shell Uniform Load Set(s) from ETABS.");
                    }
                }
                else
                {
                    string refreshErr = refreshResult != null ? refreshResult.Message : "Failed to retrieve updated definitions.";
                    SetStatus("Status: ETABS table updated, but current definitions could not be refreshed.");
                    MessageBox.Show(this, "ETABS table updated successfully, but current definitions could not be refreshed:\r\n" + refreshErr, "Refresh Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }

                ShellUniformLoadSetApplyResultDto result = applyResult.Data ?? new ShellUniformLoadSetApplyResultDto();
                string message =
                    "Shell Uniform Load Sets Updated\r\n\r\n" +
                    "Created: " + result.CreatedCount.ToString(CultureInfo.InvariantCulture) + "\r\n" +
                    "Updated: " + result.UpdatedCount.ToString(CultureInfo.InvariantCulture) + "\r\n" +
                    "Deleted: " + result.DeletedCount.ToString(CultureInfo.InvariantCulture) + "\r\n" +
                    "Load entries applied: " + result.LoadEntryCount.ToString(CultureInfo.InvariantCulture) + "\r\n\r\n" +
                    "Warnings: " + result.WarningCount.ToString(CultureInfo.InvariantCulture);
                if (!string.IsNullOrWhiteSpace(result.ImportLog) && result.WarningCount > 0)
                {
                    message += "\r\n\r\nETABS Import Log:\r\n" + result.ImportLog;
                }

                MessageBox.Show(this, message, "Shell Uniform Load Sets Updated", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            finally
            {
                Cursor.Current = previousCursor;
                btnApply.Enabled = true;
            }
        }

        private OperationResult ApplySelectedUnitSystem()
        {
            UnitOption option = cboSelectedUnits.SelectedItem as UnitOption;
            if (option == null)
            {
                return OperationResult.Failure("Please select an ETABS unit system.");
            }

            OperationResult result = _connectionService.SetPresentUnitSystem(option.ToDto());
            if (result.IsSuccess)
            {
                lblPresentUnits.Text = "Present Units: " + option.DisplayName;
                lblLoadValues.Text = "Load Values: " + option.DisplayName;
            }

            return result;
        }

        private void cboSelectedUnits_SelectedIndexChanged(object sender, EventArgs e)
        {
            UnitOption option = cboSelectedUnits.SelectedItem as UnitOption;
            if (option != null)
            {
                lblLoadValues.Text = "Load Values: " + option.DisplayName;
            }
        }

        private OperationResult<List<ShellUniformLoadSetDefinitionDto>> ValidateAndCreateDefinitions()
        {
            ClearCellHighlights();
            var definitions = new List<ShellUniformLoadSetDefinitionDto>();
            var errors = new List<string>();
            var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (DataGridViewRow row in dgvLoadSets.Rows)
            {
                if (row.IsNewRow || IsBlankGridRow(row))
                {
                    continue;
                }

                string name;
                string displayColumnKey;
                if (chkApplySuggestedName.Checked)
                {
                    name = NormalizeText(ToCellText(row.Cells[SuggestedNameColumnKey].Value));
                    displayColumnKey = SuggestedNameColumnKey;
                }
                else
                {
                    name = NormalizeText(ToCellText(row.Cells[NameColumnKey].Value));
                    displayColumnKey = NameColumnKey;
                }

                if (string.IsNullOrWhiteSpace(name))
                {
                    HighlightCell(row.Cells[displayColumnKey]);
                    errors.Add("Row " + (row.Index + 1).ToString(CultureInfo.InvariantCulture) + ": " + (chkApplySuggestedName.Checked ? "Suggested Name" : "UniformLoadSetName") + " is required.");
                    continue;
                }

                if (!names.Add(name))
                {
                    HighlightCell(row.Cells[displayColumnKey]);
                    errors.Add("Row " + (row.Index + 1).ToString(CultureInfo.InvariantCulture) + ": duplicate " + (chkApplySuggestedName.Checked ? "Suggested Name" : "UniformLoadSetName") + " '" + name + "'.");
                    continue;
                }

                var definition = new ShellUniformLoadSetDefinitionDto { Name = name };
                foreach (DataGridViewColumn column in dgvLoadSets.Columns.Cast<DataGridViewColumn>().Where(IsLoadPatternColumn))
                {
                    string text = NormalizeText(ToCellText(row.Cells[column.Index].Value));
                    if (string.IsNullOrWhiteSpace(text))
                    {
                        continue;
                    }

                    double value;
                    if (!TryParseUserNumber(text, out value))
                    {
                        HighlightCell(row.Cells[column.Index]);
                        errors.Add("Row " + (row.Index + 1).ToString(CultureInfo.InvariantCulture) + ", column " + column.HeaderText + ": value '" + text + "' is not a valid number.");
                        continue;
                    }

                    definition.LoadValuesByPattern[column.HeaderText] = value;
                }

                if (definition.LoadValuesByPattern.Count == 0)
                {
                    HighlightCell(row.Cells[NameColumnKey]);
                    errors.Add("Row " + (row.Index + 1).ToString(CultureInfo.InvariantCulture) + ": every load set must contain at least one load value.");
                    continue;
                }

                definitions.Add(definition);
            }

            return errors.Count > 0
                ? OperationResult<List<ShellUniformLoadSetDefinitionDto>>.Failure("Grid validation failed:\r\n\r\n" + string.Join("\r\n", errors))
                : OperationResult<List<ShellUniformLoadSetDefinitionDto>>.Success(definitions);
        }

        private void AddLoadPatternColumn(string patternName)
        {
            string canonicalName = NormalizeText(patternName);
            if (string.IsNullOrWhiteSpace(canonicalName) || GridHasLoadPatternColumn(canonicalName))
            {
                return;
            }

            var column = new DataGridViewTextBoxColumn
            {
                Name = canonicalName,
                HeaderText = canonicalName,
                Tag = canonicalName,
                SortMode = DataGridViewColumnSortMode.NotSortable,
                MinimumWidth = LoadPatternColumnMinimumWidth,
                Width = 110
            };
            column.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            column.HeaderCell.Style.WrapMode = DataGridViewTriState.True;
            dgvLoadSets.Columns.Add(column);
        }

        private void ConfigureGridTextWrapping()
        {
            dgvLoadSets.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
            dgvLoadSets.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            dgvLoadSets.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            dgvLoadSets.ColumnHeadersDefaultCellStyle.WrapMode = DataGridViewTriState.True;
            dgvLoadSets.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
        }

        private void AutoFitGrid()
        {
            if (dgvLoadSets.Columns.Count == 0)
            {
                return;
            }

            dgvLoadSets.SuspendLayout();
            try
            {
                dgvLoadSets.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
                foreach (DataGridViewColumn column in dgvLoadSets.Columns)
                {
                    bool isNameColumn = column.Index == 0;
                    int minimumWidth = isNameColumn ? NameColumnMinimumWidth : LoadPatternColumnMinimumWidth;
                    int maximumWidth = isNameColumn ? NameColumnMaximumWidth : LoadPatternColumnMaximumWidth;

                    column.MinimumWidth = minimumWidth;
                    column.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
                    column.HeaderCell.Style.WrapMode = DataGridViewTriState.True;

                    if (column.Width < minimumWidth)
                    {
                        column.Width = minimumWidth;
                    }
                    else if (column.Width > maximumWidth)
                    {
                        column.Width = maximumWidth;
                    }
                }

                dgvLoadSets.AutoResizeRows(DataGridViewAutoSizeRowsMode.AllCells);
                dgvLoadSets.AutoResizeColumnHeadersHeight();
            }
            finally
            {
                dgvLoadSets.ResumeLayout();
            }
        }

        private DataGridViewColumn GetSelectedLoadPatternColumn()
        {
            if (dgvLoadSets.CurrentCell != null && IsLoadPatternColumn(dgvLoadSets.CurrentCell.ColumnIndex))
            {
                return dgvLoadSets.Columns[dgvLoadSets.CurrentCell.ColumnIndex];
            }

            foreach (DataGridViewColumn column in dgvLoadSets.SelectedColumns)
            {
                if (IsLoadPatternColumn(column))
                {
                    return column;
                }
            }

            return null;
        }

        private void dgvLoadSets_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            if (e.ColumnIndex == 0)
            {
                dgvLoadSets.Columns[0].Name = NameColumnKey;
                dgvLoadSets.Columns[0].HeaderText = NameColumnKey;
            }
        }

        private void dgvLoadSets_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0 && e.ColumnIndex >= 0)
            {
                dgvLoadSets.Rows[e.RowIndex].Cells[e.ColumnIndex].Style.BackColor = Color.Empty;
                if (chkApplySuggestedName.Checked && IsLoadPatternColumn(e.ColumnIndex))
                {
                    ApplySuggestedNamesToAllRows();
                }

                AutoFitGrid();
            }
        }

        private void chkApplySuggestedName_CheckedChanged(object sender, EventArgs e)
        {
            SetSuggestedNameMode();
            if (chkApplySuggestedName.Checked)
            {
                ApplySuggestedNamesToAllRows();
                AutoFitGrid();
                SetStatus("Status: Suggested names applied.");
            }
        }

        private void SetSuggestedNameMode()
        {
            if (dgvLoadSets.Columns.Contains(SuggestedNameColumnKey))
            {
                dgvLoadSets.Columns[SuggestedNameColumnKey].Visible = chkApplySuggestedName.Checked;
            }
        }

        private void ApplySuggestedNamesIfEnabled()
        {
            if (chkApplySuggestedName.Checked)
            {
                ApplySuggestedNamesToAllRows();
            }
        }

        private void ApplySuggestedNamesToAllRows()
        {
            int suggestedIndex = 1;
            foreach (DataGridViewRow row in dgvLoadSets.Rows)
            {
                if (row.IsNewRow)
                {
                    continue;
                }

                if (!RowHasAnyLoadValue(row))
                {
                    if (dgvLoadSets.Columns.Contains(SuggestedNameColumnKey))
                    {
                        row.Cells[SuggestedNameColumnKey].Value = string.Empty;
                    }
                    continue;
                }

                if (dgvLoadSets.Columns.Contains(SuggestedNameColumnKey))
                {
                    row.Cells[SuggestedNameColumnKey].Value = CreateSuggestedName(row, suggestedIndex);
                }
                suggestedIndex++;
            }
        }

        private string CreateSuggestedName(DataGridViewRow row, int suggestedIndex)
        {
            var parts = new List<string>();
            foreach (DataGridViewColumn column in dgvLoadSets.Columns.Cast<DataGridViewColumn>().Where(IsLoadPatternColumn))
            {
                string valueText = NormalizeText(ToCellText(row.Cells[column.Index].Value));
                if (string.IsNullOrWhiteSpace(valueText))
                {
                    continue;
                }

                double value;
                string formattedValue = TryParseUserNumber(valueText, out value)
                    ? FormatSuggestedValue(value)
                    : NormalizeSuggestedNameToken(valueText);
                parts.Add(NormalizeSuggestedNameToken(column.HeaderText) + ":" + formattedValue);
            }

            return suggestedIndex.ToString("000", CultureInfo.InvariantCulture) + "-[" + string.Join("_", parts) + "]";
        }

        private void UpdateLoadPatternButtonState()
        {
            btnAddLoadPattern.Enabled = _loadPatternLookup.Count > 0 && _loadPatternLookup.Values.Any(name => !GridHasLoadPatternColumn(name));
        }

        private bool GridHasLoadPatternColumn(string patternName)
        {
            return FindLoadPatternColumn(patternName) != null;
        }

        private DataGridViewColumn FindLoadPatternColumn(string patternName)
        {
            return dgvLoadSets.Columns.Cast<DataGridViewColumn>()
                .FirstOrDefault(column => IsLoadPatternColumn(column) && string.Equals(column.HeaderText, patternName, StringComparison.OrdinalIgnoreCase));
        }

        private bool GridContainsData()
        {
            return dgvLoadSets.Columns.Cast<DataGridViewColumn>().Any(IsLoadPatternColumn) ||
                   dgvLoadSets.Rows.Cast<DataGridViewRow>().Any(row => !row.IsNewRow && !IsBlankGridRow(row));
        }

        private bool IsBlankGridRow(DataGridViewRow row)
        {
            return dgvLoadSets.Columns.Cast<DataGridViewColumn>()
                .Where(IsLoadPatternColumn)
                .All(column => string.IsNullOrWhiteSpace(ToCellText(row.Cells[column.Index].Value)));
        }

        private bool RowHasAnyLoadValue(DataGridViewRow row)
        {
            return dgvLoadSets.Columns.Cast<DataGridViewColumn>()
                .Where(IsLoadPatternColumn)
                .Any(column => !string.IsNullOrWhiteSpace(ToCellText(row.Cells[column.Index].Value)));
        }

        private bool IsBlankExcelRow(ExcelSelectedRangeData range, int row)
        {
            for (int column = 1; column <= range.ColumnCount; column++)
            {
                if (!string.IsNullOrWhiteSpace(ToCellText(range.GetValue(row, column))))
                {
                    return false;
                }
            }

            return true;
        }

        private int CountNonBlankRows()
        {
            return dgvLoadSets.Rows.Cast<DataGridViewRow>().Count(row => !row.IsNewRow && !IsBlankGridRow(row));
        }

        private void ClearCellHighlights()
        {
            foreach (DataGridViewRow row in dgvLoadSets.Rows)
            {
                foreach (DataGridViewCell cell in row.Cells)
                {
                    cell.Style.BackColor = Color.Empty;
                }
            }
        }

        private static void HighlightCell(DataGridViewCell cell)
        {
            if (cell != null)
            {
                cell.Style.BackColor = Color.LightPink;
            }
        }

        private void SetStatus(string status)
        {
            lblStatus.Text = status;
        }

        private static string ToCellText(object value)
        {
            if (value == null)
            {
                return string.Empty;
            }

            IFormattable formattable = value as IFormattable;
            if (formattable != null && !(value is string))
            {
                return formattable.ToString(null, CultureInfo.InvariantCulture);
            }

            return Convert.ToString(value, CultureInfo.CurrentCulture) ?? string.Empty;
        }

        private static string NormalizeText(string value)
        {
            return string.IsNullOrWhiteSpace(value) ? string.Empty : value.Trim();
        }

        private static bool TryParseUserNumber(string text, out double value)
        {
            string trimmed = NormalizeText(text);
            if (double.TryParse(trimmed, NumberStyles.Float, CultureInfo.CurrentCulture, out value))
            {
                return true;
            }

            if (double.TryParse(trimmed, NumberStyles.Float, CultureInfo.InvariantCulture, out value))
            {
                return true;
            }

            if (trimmed.IndexOf(',') >= 0 && trimmed.IndexOf('.') < 0)
            {
                string normalized = trimmed.Replace(',', '.');
                return double.TryParse(normalized, NumberStyles.Float, CultureInfo.InvariantCulture, out value);
            }

            value = 0;
            return false;
        }

        private static string FormatSuggestedValue(double value)
        {
            return value.ToString("G15", CultureInfo.InvariantCulture);
        }

        private static string FormatGridValue(double value)
        {
            return value.ToString("G15", CultureInfo.InvariantCulture);
        }

        private static string NormalizeSuggestedNameToken(string value)
        {
            string text = NormalizeText(value);
            if (string.IsNullOrWhiteSpace(text))
            {
                return string.Empty;
            }

            return text
                .Replace("[", string.Empty)
                .Replace("]", string.Empty)
                .Replace("\r", " ")
                .Replace("\n", " ")
                .Replace("\t", " ")
                .Trim();
        }

        private sealed class UnitOption
        {
            public UnitOption(string displayName, int forceUnit, int lengthUnit, int temperatureUnit)
            {
                DisplayName = displayName;
                ForceUnit = forceUnit;
                LengthUnit = lengthUnit;
                TemperatureUnit = temperatureUnit;
            }

            public string DisplayName { get; private set; }

            private int ForceUnit { get; set; }

            private int LengthUnit { get; set; }

            private int TemperatureUnit { get; set; }

            public bool Matches(CSISapModelPresentUnitSystemDTO unitSystem)
            {
                return unitSystem != null &&
                       unitSystem.ForceUnit == ForceUnit &&
                       unitSystem.LengthUnit == LengthUnit &&
                       unitSystem.TemperatureUnit == TemperatureUnit;
            }

            public CSISapModelPresentUnitSystemDTO ToDto()
            {
                return new CSISapModelPresentUnitSystemDTO
                {
                    ForceUnit = ForceUnit,
                    LengthUnit = LengthUnit,
                    TemperatureUnit = TemperatureUnit
                };
            }

            public override string ToString()
            {
                return DisplayName;
            }
        }
    }

    internal class Win32WindowWrapper : IWin32Window
    {
        public IntPtr Handle { get; }
        public Win32WindowWrapper(IntPtr handle)
        {
            Handle = handle;
        }
    }

    internal sealed class ProgressForm : Form
    {
        private readonly Label _lblMessage;
        private readonly ProgressBar _progressBar;

        public ProgressForm(string title, string message)
        {
            Text = title;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            StartPosition = FormStartPosition.CenterParent;
            ClientSize = new Size(400, 110);
            BackColor = Color.FromArgb(11, 31, 58); // #0B1F3A Navy background

            TableLayoutPanel panel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 2,
                Padding = new Padding(20)
            };
            panel.RowStyles.Add(new RowStyle(SizeType.Percent, 45F));
            panel.RowStyles.Add(new RowStyle(SizeType.Percent, 55F));

            _lblMessage = new Label
            {
                Text = message,
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 9.75F, FontStyle.Regular),
                TextAlign = ContentAlignment.MiddleLeft,
                Dock = DockStyle.Fill,
                AutoEllipsis = true
            };

            _progressBar = new ProgressBar
            {
                Style = ProgressBarStyle.Marquee,
                MarqueeAnimationSpeed = 30,
                Height = 18,
                Dock = DockStyle.Fill
            };

            panel.Controls.Add(_lblMessage, 0, 0);
            panel.Controls.Add(_progressBar, 0, 1);
            Controls.Add(panel);
        }
    }
}
