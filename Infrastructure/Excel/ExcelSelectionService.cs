using System;
using System.Collections.Generic;
using ExcelCSIToolBoxAddIn.Common.Results;
using Microsoft.Office.Interop.Excel;

namespace ExcelCSIToolBoxAddIn.Infrastructure.Excel
{
    public class ExcelSelectionService : IExcelSelectionService
    {
        public OperationResult<IReadOnlyList<string>> ReadSingleColumnTextValues()
        {
            var selectionResult = GetActiveSelection();
            if (!selectionResult.IsSuccess)
            {
                return OperationResult<IReadOnlyList<string>>.Failure(selectionResult.Message);
            }

            var selection = selectionResult.Data;
            if (selection.Columns.Count != 1)
            {
                return OperationResult<IReadOnlyList<string>>.Failure("Please select exactly 1 column and N rows.");
            }

            var values = new List<string>();
            int rowCount = selection.Rows.Count;

            for (int row = 1; row <= rowCount; row++)
            {
                string value = ReadCellText(selection, row, 1);
                if (!string.IsNullOrWhiteSpace(value))
                {
                    values.Add(value);
                }
            }

            if (values.Count == 0)
            {
                return OperationResult<IReadOnlyList<string>>.Failure("The selected Excel range does not contain any non-empty values.");
            }

            return OperationResult<IReadOnlyList<string>>.Success(values);
        }

        public OperationResult<IReadOnlyList<ExcelFrameByPointRow>> ReadFrameByPointRows()
        {
            var selectionResult = GetActiveSelection();
            if (!selectionResult.IsSuccess)
            {
                return OperationResult<IReadOnlyList<ExcelFrameByPointRow>>.Failure(selectionResult.Message);
            }

            var selection = selectionResult.Data;
            if (selection.Columns.Count != 4)
            {
                return OperationResult<IReadOnlyList<ExcelFrameByPointRow>>.Failure("Please select exactly 4 columns in this order: UniqueName, Section, PointIName, PointJName.");
            }

            var rows = new List<ExcelFrameByPointRow>();
            int rowCount = selection.Rows.Count;

            for (int row = 1; row <= rowCount; row++)
            {
                rows.Add(new ExcelFrameByPointRow
                {
                    ExcelRowNumber = selection.Row + row - 1,
                    UniqueNameText = ReadCellText(selection, row, 1),
                    SectionText = ReadCellText(selection, row, 2),
                    PointINameText = ReadCellText(selection, row, 3),
                    PointJNameText = ReadCellText(selection, row, 4)
                });
            }

            if (rows.Count == 0)
            {
                return OperationResult<IReadOnlyList<ExcelFrameByPointRow>>.Failure("Please select at least one row.");
            }

            return OperationResult<IReadOnlyList<ExcelFrameByPointRow>>.Success(rows);
        }

        private static OperationResult<Range> GetActiveSelection()
        {
            try
            {
                Application excelApp = Globals.ExcelCSIToolBoxAddin.Application;
                if (excelApp == null)
                {
                    return OperationResult<Range>.Failure("Excel application is not available.");
                }

                var selection = excelApp.Selection as Range;
                if (selection == null)
                {
                    return OperationResult<Range>.Failure("Please select a range in Excel and try again.");
                }

                return OperationResult<Range>.Success(selection);
            }
            catch (Exception)
            {
                return OperationResult<Range>.Failure("Unable to read the current Excel selection.");
            }
        }

        private static string ReadCellText(Range selection, int row, int column)
        {
            try
            {
                var cell = selection.Cells[row, column] as Range;
                if (cell == null || cell.Value2 == null)
                {
                    return null;
                }

                return Convert.ToString(cell.Value2)?.Trim();
            }
            catch
            {
                return null;
            }
        }
    }
}
