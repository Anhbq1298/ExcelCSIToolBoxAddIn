using System;
using System.Runtime.InteropServices;
using ExcelCSIToolBox.Core.Common.Results;
using ExcelCSIToolBox.Infrastructure.Excel.Interop;
using Microsoft.Office.Interop.Excel;

namespace ExcelCSIToolBoxAddIn.UI.Forms
{
    internal sealed class ExcelSelectedRangeReader
    {
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool IsIconic(IntPtr hWnd);

        private const int SW_RESTORE = 9;
        private const int SW_SHOW = 5;

        public void ActivateExcel()
        {
            try
            {
                Application application = ExcelApplicationProvider.GetApplication();
                IntPtr hwnd = IntPtr.Zero;
                if (application != null)
                {
                    hwnd = new IntPtr(application.Hwnd);
                }

                if (hwnd == IntPtr.Zero)
                {
                    hwnd = System.Diagnostics.Process.GetCurrentProcess().MainWindowHandle;
                }

                if (hwnd != IntPtr.Zero)
                {
                    if (IsIconic(hwnd))
                    {
                        ShowWindow(hwnd, SW_RESTORE);
                    }
                    else
                    {
                        ShowWindow(hwnd, SW_SHOW);
                    }
                    SetForegroundWindow(hwnd);

                    if (application != null && application.ActiveWindow != null)
                    {
                        application.ActiveWindow.Activate();
                    }
                }
            }
            catch
            {
                // Ignore activation failures
            }
        }

        public OperationResult<ExcelSelectedRangeData> ReadSelectedRange()
        {
            try
            {
                Application application = ExcelApplicationProvider.GetApplication();
                if (application == null)
                {
                    return OperationResult<ExcelSelectedRangeData>.Failure("Excel application is not available.");
                }

                ActivateExcel();
                System.Windows.Forms.Application.DoEvents();
                System.Threading.Thread.Sleep(100);

                object input = application.InputBox(
                    "Select a Shell Uniform Load Set range:\r\nUniformLoadSetName | load pattern columns...",
                    "Select Shell Uniform Load Set Range",
                    Type: 8);
                if (input is bool canceled && !canceled)
                {
                    return OperationResult<ExcelSelectedRangeData>.Failure("Action canceled by user.");
                }

                Range selectedRange = input as Range;
                if (selectedRange == null)
                {
                    return OperationResult<ExcelSelectedRangeData>.Failure("Please select a rectangular Excel range and try again.");
                }

                int rowCount = selectedRange.Rows.Count;
                int columnCount = selectedRange.Columns.Count;
                if (rowCount < 1 || columnCount < 1)
                {
                    return OperationResult<ExcelSelectedRangeData>.Failure("The selected Excel range is empty.");
                }

                object rawValue = selectedRange.Value2;
                object[,] values = new object[rowCount + 1, columnCount + 1];
                object[,] matrix = rawValue as object[,];
                if (matrix != null)
                {
                    for (int row = 1; row <= rowCount; row++)
                    {
                        for (int column = 1; column <= columnCount; column++)
                        {
                            object value = matrix[row, column];
                            if (IsExcelErrorValue(value))
                            {
                                return OperationResult<ExcelSelectedRangeData>.Failure("Cannot import the selected Excel range.\r\n\r\nExcel error values such as #VALUE! or #N/A are not supported.");
                            }

                            values[row, column] = value;
                        }
                    }
                }
                else
                {
                    if (IsExcelErrorValue(rawValue))
                    {
                        return OperationResult<ExcelSelectedRangeData>.Failure("Cannot import the selected Excel range.\r\n\r\nExcel error values such as #VALUE! or #N/A are not supported.");
                    }

                    values[1, 1] = rawValue;
                }

                return OperationResult<ExcelSelectedRangeData>.Success(new ExcelSelectedRangeData(values, rowCount, columnCount));
            }
            catch (Exception ex)
            {
                return OperationResult<ExcelSelectedRangeData>.Failure("Unable to read the current Excel selection: " + ex.Message);
            }
        }

        private static bool IsExcelErrorValue(object value)
        {
            return value is ErrorWrapper;
        }
    }
}
