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

            try
            {
                Microsoft.Office.Interop.Excel.Application excelApp = ExcelApplicationProvider.GetApplication();
                if (excelApp == null)
                {
                    return OperationResult.Failure("Excel application is not available.");
                }

                Range activeCell = excelApp.ActiveCell;
                if (activeCell == null)
                {
                    return OperationResult.Failure("Please select a target cell in Excel and try again.");
                }

                Worksheet sheet = activeCell.Worksheet;
                int startRow = activeCell.Row;
                int startCol = activeCell.Column;

                for (int col = 0; col < dataFrame.Columns.Count; col++)
                {
                    sheet.Cells[startRow, startCol + col] = dataFrame.Columns[col];
                }

                for (int rowIndex = 0; rowIndex < dataFrame.Rows.Count; rowIndex++)
                {
                    var row = dataFrame.Rows[rowIndex];

                    if (row == null)
                    {
                        continue;
                    }

                    int maxCol = row.Count < dataFrame.Columns.Count ? row.Count : dataFrame.Columns.Count;

                    for (int col = 0; col < maxCol; col++)
                    {
                        sheet.Cells[startRow + rowIndex + 1, startCol + col] = row[col];
                    }
                }

                return OperationResult.Success($"Successfully exported {dataFrame.Rows.Count} row(s) to Excel.");
            }
            catch (Exception)
            {
                return OperationResult.Failure("Failed to write table data to Excel.");
            }
        }
    }
}

