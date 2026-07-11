using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;
using ExcelCSIToolBox.Core.Abstractions.CSI;
using ExcelCSIToolBox.Core.Common.Results;
using ExcelCSIToolBox.Data.DTOs.CSI;

namespace ExcelCSIToolBoxAddIn.UI.Forms
{
    internal partial class ShellUniformLoadSetForm : Form
    {
        private const string NameColumnKey = "UniformLoadSetName";
        private readonly ICSISapModelConnectionService _connectionService;
        private readonly ExcelSelectedRangeReader _excelRangeReader;
        private readonly Dictionary<string, string> _loadPatternLookup = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        public ShellUniformLoadSetForm(ICSISapModelConnectionService connectionService)
        {
            if (connectionService == null) throw new ArgumentNullException(nameof(connectionService));
            _connectionService = connectionService;
            _excelRangeReader = new ExcelSelectedRangeReader();
            InitializeComponent();
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
                return;
            }

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

            UpdateLoadPatternButtonState();
            SetStatus("Status: Ready.");
        }

        private void InitializeGrid()
        {
            dgvLoadSets.Columns.Clear();
            dgvLoadSets.Rows.Clear();
            DataGridViewTextBoxColumn nameColumn = new DataGridViewTextBoxColumn
            {
                Name = NameColumnKey,
                HeaderText = NameColumnKey,
                Tag = NameColumnKey,
                SortMode = DataGridViewColumnSortMode.NotSortable,
                Width = 220,
                Frozen = true
            };
            dgvLoadSets.Columns.Add(nameColumn);
        }

        private void btnImportExcelRange_Click(object sender, EventArgs e)
        {
            OperationResult<ExcelSelectedRangeData> readResult = _excelRangeReader.ReadSelectedRange();
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

            if (GridContainsData())
            {
                DialogResult replace = MessageBox.Show(
                    this,
                    "The current table contains data.\r\n\r\nImporting the selected Excel range will replace the current table.\r\n\r\nContinue?",
                    "Replace Current Table",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);
                if (replace != DialogResult.Yes)
                {
                    return;
                }
            }

            TransferExcelDataToGrid(range, headerResult.Data);
            SetStatus("Status: " + CountNonBlankRows().ToString(CultureInfo.InvariantCulture) + " load sets ready.");
        }

        private OperationResult<List<string>> ValidateImportHeaders(ExcelSelectedRangeData range)
        {
            string firstHeader = ToCellText(range.GetValue(1, 1));
            if (!string.Equals(NormalizeText(firstHeader), NameColumnKey, StringComparison.OrdinalIgnoreCase))
            {
                return OperationResult<List<string>>.Failure(
                    "Cannot import the selected Excel range.\r\n\r\n" +
                    "The first header must be \"UniformLoadSetName\".\r\n\r\n" +
                    "Current header: \"" + (firstHeader ?? string.Empty) + "\"");
            }

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
            InitializeGrid();
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
                for (int column = 1; column <= range.ColumnCount; column++)
                {
                    gridRow.Cells[column - 1].Value = ToCellText(range.GetValue(row, column));
                }
            }

            UpdateLoadPatternButtonState();
        }

        private void btnAddRow_Click(object sender, EventArgs e)
        {
            dgvLoadSets.Rows.Add();
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
            UpdateLoadPatternButtonState();
            SetStatus("Status: Ready.");
        }

        private void btnApply_Click(object sender, EventArgs e)
        {
            OperationResult<List<ShellUniformLoadSetDefinitionDto>> parseResult = ValidateAndCreateDefinitions();
            if (!parseResult.IsSuccess)
            {
                SetStatus("Status: Fix highlighted cells and try again.");
                MessageBox.Show(this, parseResult.Message, "Grid Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            btnApply.Enabled = false;
            Cursor previousCursor = Cursor.Current;
            Cursor.Current = Cursors.WaitCursor;
            try
            {
                OperationResult<ShellUniformLoadSetApplyResultDto> applyResult = _connectionService.ApplyShellUniformLoadSets(parseResult.Data);
                if (!applyResult.IsSuccess)
                {
                    SetStatus("Status: ETABS table update failed.");
                    MessageBox.Show(this, applyResult.Message, "ETABS Import Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                ShellUniformLoadSetApplyResultDto result = applyResult.Data ?? new ShellUniformLoadSetApplyResultDto();
                string message =
                    "Shell Uniform Load Sets Updated\r\n\r\n" +
                    "Created: " + result.CreatedCount.ToString(CultureInfo.InvariantCulture) + "\r\n" +
                    "Updated: " + result.UpdatedCount.ToString(CultureInfo.InvariantCulture) + "\r\n" +
                    "Load entries applied: " + result.LoadEntryCount.ToString(CultureInfo.InvariantCulture) + "\r\n\r\n" +
                    "Warnings: " + result.WarningCount.ToString(CultureInfo.InvariantCulture);
                if (!string.IsNullOrWhiteSpace(result.ImportLog) && result.WarningCount > 0)
                {
                    message += "\r\n\r\nETABS Import Log:\r\n" + result.ImportLog;
                }

                SetStatus("Status: ETABS table updated successfully.");
                MessageBox.Show(this, message, "Shell Uniform Load Sets Updated", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            finally
            {
                Cursor.Current = previousCursor;
                btnApply.Enabled = true;
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

                string name = NormalizeText(ToCellText(row.Cells[NameColumnKey].Value));
                if (string.IsNullOrWhiteSpace(name))
                {
                    HighlightCell(row.Cells[NameColumnKey]);
                    errors.Add("Row " + (row.Index + 1).ToString(CultureInfo.InvariantCulture) + ": UniformLoadSetName is required.");
                    continue;
                }

                if (!names.Add(name))
                {
                    HighlightCell(row.Cells[NameColumnKey]);
                    errors.Add("Row " + (row.Index + 1).ToString(CultureInfo.InvariantCulture) + ": duplicate UniformLoadSetName '" + name + "'.");
                    continue;
                }

                var definition = new ShellUniformLoadSetDefinitionDto { Name = name };
                foreach (DataGridViewColumn column in dgvLoadSets.Columns.Cast<DataGridViewColumn>().Where(column => column.Index > 0))
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

            if (definitions.Count == 0)
            {
                errors.Add("At least one load set row is required.");
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
                Width = 110
            };
            dgvLoadSets.Columns.Add(column);
        }

        private DataGridViewColumn GetSelectedLoadPatternColumn()
        {
            if (dgvLoadSets.CurrentCell != null && dgvLoadSets.CurrentCell.ColumnIndex > 0)
            {
                return dgvLoadSets.Columns[dgvLoadSets.CurrentCell.ColumnIndex];
            }

            foreach (DataGridViewColumn column in dgvLoadSets.SelectedColumns)
            {
                if (column.Index > 0)
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
            }
        }

        private void UpdateLoadPatternButtonState()
        {
            btnAddLoadPattern.Enabled = _loadPatternLookup.Count > 0 && _loadPatternLookup.Values.Any(name => !GridHasLoadPatternColumn(name));
        }

        private bool GridHasLoadPatternColumn(string patternName)
        {
            return dgvLoadSets.Columns.Cast<DataGridViewColumn>()
                .Any(column => column.Index > 0 && string.Equals(column.HeaderText, patternName, StringComparison.OrdinalIgnoreCase));
        }

        private bool GridContainsData()
        {
            return dgvLoadSets.Columns.Count > 1 || dgvLoadSets.Rows.Cast<DataGridViewRow>().Any(row => !row.IsNewRow && !IsBlankGridRow(row));
        }

        private bool IsBlankGridRow(DataGridViewRow row)
        {
            return dgvLoadSets.Columns.Cast<DataGridViewColumn>()
                .All(column => string.IsNullOrWhiteSpace(ToCellText(row.Cells[column.Index].Value)));
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
    }
}
