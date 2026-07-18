using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using ExcelCSIToolBox.Application.Modelling.DropPanels;
using ExcelCSIToolBox.Core.Common.Results;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelCSIToolBoxAddIn.UI.Views
{
    public sealed class DropPanelExcelLogExporter
    {
        private const string WorksheetName = "DROP_PANEL_LOG";

        public OperationResult Export(IReadOnlyList<DropPanelLogEntry> entries)
        {
            if (entries == null || entries.Count == 0)
            {
                return OperationResult.Failure("There are no Drop Panel log entries to export.");
            }

            Excel.Application application = Globals.ExcelCSIToolBoxAddin == null
                ? null
                : Globals.ExcelCSIToolBoxAddin.Application;
            if (application == null)
            {
                return OperationResult.Failure("No active Excel workbook is available.");
            }

            Excel.Workbook workbook = application.ActiveWorkbook;
            if (workbook == null)
            {
                return OperationResult.Failure("No active Excel workbook is available.");
            }

            bool screenUpdating = application.ScreenUpdating;
            bool enableEvents = application.EnableEvents;
            bool displayAlerts = application.DisplayAlerts;
            Excel.XlCalculation calculation = application.Calculation;
            Excel.Worksheet activeWorksheet = application.ActiveSheet as Excel.Worksheet;
            Excel.Sheets worksheets = null;
            Excel.Worksheet worksheet = null;
            Excel.Range allCells = null;
            Excel.Range startCell = null;
            Excel.Range targetRange = null;
            Excel.Range rows = null;
            Excel.Range columns = null;

            try
            {
                application.ScreenUpdating = false;
                application.EnableEvents = false;
                application.DisplayAlerts = false;
                application.Calculation = Excel.XlCalculation.xlCalculationManual;

                worksheets = workbook.Worksheets;
                for (int index = 1; index <= worksheets.Count; index++)
                {
                    Excel.Worksheet candidate = worksheets[index] as Excel.Worksheet;
                    if (candidate != null && string.Equals(candidate.Name, WorksheetName, StringComparison.OrdinalIgnoreCase))
                    {
                        worksheet = candidate;
                        break;
                    }

                    Release(candidate);
                }

                if (worksheet == null)
                {
                    worksheet = worksheets.Add() as Excel.Worksheet;
                    if (worksheet == null)
                    {
                        return OperationResult.Failure("Excel could not create the DROP_PANEL_LOG worksheet.");
                    }

                    worksheet.Name = WorksheetName;
                }

                allCells = worksheet.Cells;
                allCells.ClearContents();
                object[,] values = BuildValues(entries);
                startCell = allCells[1, 1] as Excel.Range;
                targetRange = startCell.Resize[values.GetLength(0), values.GetLength(1)];
                targetRange.Value2 = values;

                rows = targetRange.Rows;
                Excel.Range header = rows[1] as Excel.Range;
                if (header != null)
                {
                    Excel.Font headerFont = null;
                    try
                    {
                        headerFont = header.Font;
                        headerFont.Bold = true;
                        header.WrapText = true;
                    }
                    finally
                    {
                        Release(headerFont);
                        Release(header);
                    }
                }

                columns = targetRange.Columns;
                columns.AutoFit();
                worksheet.Activate();
                return OperationResult.Success("Exported " + entries.Count + " Drop Panel log row(s) to " + WorksheetName + ".");
            }
            catch (Exception ex)
            {
                return OperationResult.Failure("Failed to export the Drop Panel log to Excel: " + ex.Message);
            }
            finally
            {
                try
                {
                    application.Calculation = calculation;
                    application.DisplayAlerts = displayAlerts;
                    application.EnableEvents = enableEvents;
                    application.ScreenUpdating = screenUpdating;
                    if (activeWorksheet != null)
                    {
                        activeWorksheet.Activate();
                    }
                }
                finally
                {
                    Release(columns);
                    Release(rows);
                    Release(targetRange);
                    Release(startCell);
                    Release(allCells);
                    Release(worksheet);
                    Release(worksheets);
                    Release(activeWorksheet);
                }
            }
        }

        private static object[,] BuildValues(IReadOnlyList<DropPanelLogEntry> entries)
        {
            string[] headers =
            {
                "Timestamp", "ETABS Model", "Story", "Column", "Source Area", "New Area", "Region Type",
                "Original Property", "New Property", "Direct Load Status", "Shell Load Set Status",
                "Local Axis Status", "Local 3 Status", "Diaphragm Status", "Verification Status", "Message"
            };
            object[,] values = new object[entries.Count + 1, headers.Length];
            for (int column = 0; column < headers.Length; column++)
            {
                values[0, column] = headers[column];
            }

            for (int row = 0; row < entries.Count; row++)
            {
                DropPanelLogEntry entry = entries[row];
                values[row + 1, 0] = entry.Timestamp;
                values[row + 1, 1] = entry.EtabsModel;
                values[row + 1, 2] = entry.Story;
                values[row + 1, 3] = entry.Column;
                values[row + 1, 4] = entry.SourceArea;
                values[row + 1, 5] = entry.NewArea;
                values[row + 1, 6] = entry.RegionType;
                values[row + 1, 7] = entry.OriginalProperty;
                values[row + 1, 8] = entry.NewProperty;
                values[row + 1, 9] = entry.DirectLoadStatus;
                values[row + 1, 10] = entry.ShellLoadSetStatus;
                values[row + 1, 11] = entry.LocalAxisStatus;
                values[row + 1, 12] = entry.Local3Status;
                values[row + 1, 13] = entry.DiaphragmStatus;
                values[row + 1, 14] = entry.VerificationStatus;
                values[row + 1, 15] = entry.Message;
            }

            return values;
        }

        private static void Release(object comObject)
        {
            if (comObject != null && Marshal.IsComObject(comObject))
            {
                Marshal.ReleaseComObject(comObject);
            }
        }
    }
}
