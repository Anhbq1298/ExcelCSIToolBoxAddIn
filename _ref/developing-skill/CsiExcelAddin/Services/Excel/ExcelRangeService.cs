using System;
using System.Collections.Generic;
using CsiExcelAddin.Services.Interfaces;
using Microsoft.Office.Interop.Excel;

namespace CsiExcelAddin.Services.Excel
{
    /// <summary>
    /// Reads data from Excel ranges using the Excel Interop API.
    /// All Excel interop concerns are contained in this class —
    /// ViewModels and CSI services never touch Excel objects directly.
    /// </summary>
    public class ExcelRangeReader : IExcelRangeReader
    {
        private readonly Application _excelApp;

        /// <param name="excelApp">
        /// The running Excel Application instance.
        /// In VSTO add-ins this is available via Globals.ThisAddIn.Application.
        /// </param>
        public ExcelRangeReader(Application excelApp)
        {
            _excelApp = excelApp ?? throw new ArgumentNullException(nameof(excelApp));
        }

        public IReadOnlyList<IReadOnlyList<string>> ReadSelectedRange()
        {
            var selection = _excelApp.Selection as Range;
            if (selection == null)
                throw new InvalidOperationException("No range is currently selected in Excel.");

            return ReadRange(selection);
        }

        public IReadOnlyList<IReadOnlyList<string>> ReadNamedRange(string rangeName)
        {
            if (string.IsNullOrWhiteSpace(rangeName))
                throw new ArgumentNullException(nameof(rangeName));

            Workbook wb = _excelApp.ActiveWorkbook
                ?? throw new InvalidOperationException("No workbook is open.");

            Range range;
            try
            {
                range = wb.Names[rangeName].RefersToRange;
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException(
                    $"Named range '{rangeName}' was not found in the active workbook.", ex);
            }

            return ReadRange(range);
        }

        /// <summary>
        /// Reads all cells in the given range into a row-major list of string values.
        /// Empty rows at the tail of the selection are excluded.
        /// </summary>
        private static IReadOnlyList<IReadOnlyList<string>> ReadRange(Range range)
        {
            object[,] values = range.Value2 as object[,];
            if (values == null) return Array.Empty<IReadOnlyList<string>>();

            int rowCount = values.GetLength(0);
            int colCount = values.GetLength(1);
            var result = new List<IReadOnlyList<string>>(rowCount);

            for (int r = 1; r <= rowCount; r++)
            {
                var row = new List<string>(colCount);
                bool hasData = false;

                for (int c = 1; c <= colCount; c++)
                {
                    string cellValue = values[r, c]?.ToString() ?? string.Empty;
                    row.Add(cellValue);
                    if (!string.IsNullOrWhiteSpace(cellValue)) hasData = true;
                }

                // Skip entirely blank rows to avoid importing trailing empty rows
                if (hasData) result.Add(row);
            }

            return result;
        }
    }

    /// <summary>
    /// Writes data to Excel ranges using the Excel Interop API.
    /// </summary>
    public class ExcelRangeWriter : IExcelRangeWriter
    {
        private readonly Application _excelApp;

        public ExcelRangeWriter(Application excelApp)
        {
            _excelApp = excelApp ?? throw new ArgumentNullException(nameof(excelApp));
        }

        public void WriteToNamedRange(string rangeName, IReadOnlyList<IReadOnlyList<object>> data)
        {
            var topLeft = GetTopLeftOfNamedRange(rangeName);
            WriteFromCell(topLeft, data);
        }

        public void WriteToCell(string cellAddress, IReadOnlyList<IReadOnlyList<object>> data)
        {
            Worksheet ws = _excelApp.ActiveSheet as Worksheet
                ?? throw new InvalidOperationException("Active sheet is not a worksheet.");

            Range startCell = ws.Range[cellAddress];
            WriteFromCell(startCell, data);
        }

        public void ClearNamedRange(string rangeName)
        {
            var range = GetNamedRange(rangeName);
            range.ClearContents();
        }

        // ── Internals ─────────────────────────────────────────────────────────

        private void WriteFromCell(Range startCell, IReadOnlyList<IReadOnlyList<object>> data)
        {
            if (data == null || data.Count == 0) return;

            int rowCount = data.Count;
            int colCount = 0;
            foreach (var row in data)
                if (row.Count > colCount) colCount = row.Count;

            // Build a 2-D array — one bulk assignment is far faster than cell-by-cell writes
            object[,] values = new object[rowCount, colCount];
            for (int r = 0; r < rowCount; r++)
                for (int c = 0; c < data[r].Count; c++)
                    values[r, c] = data[r][c];

            Range target = startCell.Resize[rowCount, colCount];
            target.Value2 = values;
        }

        private Range GetTopLeftOfNamedRange(string rangeName)
        {
            var range = GetNamedRange(rangeName);
            return range.Cells[1, 1] as Range;
        }

        private Range GetNamedRange(string rangeName)
        {
            Workbook wb = _excelApp.ActiveWorkbook
                ?? throw new InvalidOperationException("No workbook is open.");
            try
            {
                return wb.Names[rangeName].RefersToRange;
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException(
                    $"Named range '{rangeName}' was not found in the active workbook.", ex);
            }
        }
    }
}
