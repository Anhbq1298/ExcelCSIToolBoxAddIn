using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

namespace ExcelCSIToolBox.Infrastructure.Excel
{
    public static class ExcelApplicationProvider
    {
        public static Func<Application> Current { get; set; }

        public static Application GetApplication()
        {
            if (Current != null)
            {
                return Current();
            }

            return Marshal.GetActiveObject("Excel.Application") as Application;
        }
    }
}
