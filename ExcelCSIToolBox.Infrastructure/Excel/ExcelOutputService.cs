using System;
using ExcelCSIToolBox.Core.Tabular;
using ExcelCSIToolBox.Core.Common.Results;
using ExcelCSIToolBox.Core.Abstractions.Excel;
using Microsoft.Office.Interop.Excel;

namespace ExcelCSIToolBox.Infrastructure.Excel
{
    public class ExcelOutputService : IExcelOutputService
    {
        public OperationResult WriteDataFrameToActiveCell(DataFrame dataFrame)
        {
            if (dataFrame == null || dataFrame.Columns == null || dataFrame.Columns.Count == 0)
            {
                return OperationResult.Failure("There is no tabular data to export.");
            }

            object[,] values = CreateValues(dataFrame);
            return WriteValuesToActiveCell(values, $"Successfully exported {dataFrame.Rows.Count} row(s) to Excel.");
        }

        public OperationResult WriteValuesToActiveCell(object[,] values, string successMessage = null)
        {
            if (values == null || values.GetLength(0) == 0 || values.GetLength(1) == 0)
            {
                return OperationResult.Failure("There is no tabular data to export.");
            }

            try
            {
                Microsoft.Office.Interop.Excel.Application excelApp = ExcelApplicationProvider.GetApplication();
                if (excelApp == null)
                {
                    return OperationResult.Failure("Excel application is not available.");
                }

                Range startCell = GetTopLeftSelectedCell(excelApp);
                if (startCell == null)
                {
                    return OperationResult.Failure("Please select a target cell in Excel and try again.");
                }

                return WriteValuesToRange(values, startCell, successMessage);
            }
            catch (Exception)
            {
                return OperationResult.Failure("Failed to write table data to Excel.");
            }
        }

        public OperationResult WriteValuesToSelectedCell(object[,] values, string prompt, string title, string successMessage = null)
        {
            if (values == null || values.GetLength(0) == 0 || values.GetLength(1) == 0)
            {
                return OperationResult.Failure("There is no tabular data to export.");
            }

            try
            {
                Microsoft.Office.Interop.Excel.Application excelApp = ExcelApplicationProvider.GetApplication();
                if (excelApp == null)
                {
                    return OperationResult.Failure("Excel application is not available.");
                }

                Range startCell = GetPromptedTopLeftCell(excelApp, prompt, title);
                if (startCell == null)
                {
                    return OperationResult.Failure("Export canceled. No target cell was selected.");
                }

                return WriteValuesToRange(values, startCell, successMessage);
            }
            catch (Exception)
            {
                return OperationResult.Failure("Failed to write table data to Excel.");
            }
        }

        private static object[,] CreateValues(DataFrame dataFrame)
        {
            int rowCount = dataFrame.Rows.Count + 1;
            int columnCount = dataFrame.Columns.Count;
            var values = new object[rowCount, columnCount];

            for (int col = 0; col < columnCount; col++)
            {
                values[0, col] = dataFrame.Columns[col];
            }

            for (int rowIndex = 0; rowIndex < dataFrame.Rows.Count; rowIndex++)
            {
                var row = dataFrame.Rows[rowIndex];
                if (row == null)
                {
                    continue;
                }

                int maxCol = row.Count < columnCount ? row.Count : columnCount;
                for (int col = 0; col < maxCol; col++)
                {
                    values[rowIndex + 1, col] = row[col];
                }
            }

            return values;
        }

        private static Range GetTopLeftSelectedCell(Microsoft.Office.Interop.Excel.Application excelApp)
        {
            var selectedRange = excelApp.Selection as Range;
            if (selectedRange != null)
            {
                return selectedRange.Cells[1, 1] as Range;
            }

            return excelApp.ActiveCell;
        }

        private static Range GetPromptedTopLeftCell(Microsoft.Office.Interop.Excel.Application excelApp, string prompt, string title)
        {
            object result = excelApp.InputBox(
                string.IsNullOrWhiteSpace(prompt) ? "Select the top-left target cell for export:" : prompt,
                string.IsNullOrWhiteSpace(title) ? "Select Export Target" : title,
                Type: 8);

            if (result is bool b && !b)
            {
                return null;
            }

            var selectedRange = result as Range;
            return selectedRange == null ? null : selectedRange.Cells[1, 1] as Range;
        }

        private static OperationResult WriteValuesToRange(object[,] values, Range startCell, string successMessage)
        {
            int rowCount = values.GetLength(0);
            int columnCount = values.GetLength(1);
            Range targetRange = startCell.Resize[rowCount, columnCount];
            targetRange.Value2 = values;

            return OperationResult.Success(successMessage ?? $"Successfully exported {rowCount - 1} row(s) to Excel.");
        }
    }
}

