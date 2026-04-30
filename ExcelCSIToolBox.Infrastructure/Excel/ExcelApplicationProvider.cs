using System;
using System.Runtime.InteropServices;
using ExcelApplication = Microsoft.Office.Interop.Excel.Application;

namespace ExcelCSIToolBox.Infrastructure.Excel
{
    public static class ExcelApplicationProvider
    {
        public static Func<ExcelApplication> Current { get; set; }

        public static ExcelApplication GetApplication()
        {
            if (Current != null)
            {
                return Current();
            }

            return Marshal.GetActiveObject("Excel.Application") as ExcelApplication;
        }
    }
}
