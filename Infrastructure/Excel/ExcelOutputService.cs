using System;
using ExcelCSIToolBoxAddIn.Common.Results;
using ExcelCSIToolBoxAddIn.Infrastructure.Etabs;
using Microsoft.Office.Interop.Excel;

namespace ExcelCSIToolBoxAddIn.Infrastructure.Excel
{
    public class ExcelOutputService : IExcelOutputService
    {
        public OperationResult WritePointsToActiveCell(System.Collections.Generic.IReadOnlyList<EtabsPointData> points)
        {
            if (points == null || points.Count == 0)
            {
                return OperationResult.Failure("There are no points to export.");
            }

            try
            {
                Application excelApp = Globals.ExcelCSIToolBoxAddin.Application;
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

                sheet.Cells[startRow, startCol] = "PointUniqueName";
                sheet.Cells[startRow, startCol + 1] = "X";
                sheet.Cells[startRow, startCol + 2] = "Y";
                sheet.Cells[startRow, startCol + 3] = "Z";

                for (int i = 0; i < points.Count; i++)
                {
                    int row = startRow + i + 1;
                    sheet.Cells[row, startCol] = points[i].PointUniqueName;
                    sheet.Cells[row, startCol + 1] = points[i].X;
                    sheet.Cells[row, startCol + 2] = points[i].Y;
                    sheet.Cells[row, startCol + 3] = points[i].Z;
                }

                return OperationResult.Success();
            }
            catch (Exception)
            {
                return OperationResult.Failure("Failed to write selected points to Excel.");
            }
        }
    }
}
