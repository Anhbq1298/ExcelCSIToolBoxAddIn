using System;
using ExcelCSIToolBox.Application.Interfaces.Excel;
using ExcelCSIToolBox.Core.Models.AnalysisResults;
using ExcelCSIToolBox.Infrastructure.Excel;
using InteropExcel = Microsoft.Office.Interop.Excel;

namespace ExcelCSIToolBox.Infrastructure.Services.Excel
{
    public class ExcelExportService : IExcelExportService
    {
        public void ExportTable(EtabsTableResult result)
        {
            if (result == null)
            {
                return;
            }

            InteropExcel.Application excelApp = ExcelApplicationProvider.GetApplication();
            if (excelApp == null)
            {
                throw new InvalidOperationException("Excel application is not available.");
            }

            InteropExcel.Workbook workbook = excelApp.ActiveWorkbook ?? excelApp.Workbooks.Add();
            InteropExcel.Worksheet worksheet = workbook.Worksheets.Add() as InteropExcel.Worksheet;
            if (worksheet == null)
            {
                throw new InvalidOperationException("Failed to create an Excel worksheet.");
            }

            worksheet.Name = GetUniqueSheetName(workbook, result.TableName);

            for (int col = 0; col < result.Headers.Count; col++)
            {
                worksheet.Cells[1, col + 1] = result.Headers[col];
            }

            for (int row = 0; row < result.Rows.Count; row++)
            {
                for (int col = 0; col < result.Rows[row].Count; col++)
                {
                    worksheet.Cells[row + 2, col + 1] = result.Rows[row][col];
                }
            }

            InteropExcel.Range usedRange = worksheet.UsedRange;
            usedRange.Columns.AutoFit();
        }

        private static string GetUniqueSheetName(InteropExcel.Workbook workbook, string name)
        {
            string baseName = GetSafeSheetName(name);
            string candidate = baseName;
            int suffix = 2;

            while (WorksheetNameExists(workbook, candidate))
            {
                string suffixText = " " + suffix;
                int maxBaseLength = 31 - suffixText.Length;
                candidate = (baseName.Length > maxBaseLength ? baseName.Substring(0, maxBaseLength) : baseName) + suffixText;
                suffix++;
            }

            return candidate;
        }

        private static bool WorksheetNameExists(InteropExcel.Workbook workbook, string name)
        {
            foreach (InteropExcel.Worksheet sheet in workbook.Worksheets)
            {
                if (string.Equals(sheet.Name, name, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }

            return false;
        }

        private static string GetSafeSheetName(string name)
        {
            string safeName = string.IsNullOrWhiteSpace(name) ? "ETABS Result" : name;

            foreach (char c in new[] { '\\', '/', '?', '*', '[', ']', ':' })
            {
                safeName = safeName.Replace(c.ToString(), string.Empty);
            }

            if (safeName.Length > 31)
            {
                safeName = safeName.Substring(0, 31);
            }

            return string.IsNullOrWhiteSpace(safeName) ? "ETABS Result" : safeName;
        }
    }
}
