using System;
using System.IO;
using System.Windows.Forms;

namespace ExcelCSIToolBoxAddIn.AddIn.Diagnostics
{
    internal static class AddInDiagnostics
    {
        private const string FolderName = "ExcelCSIToolBoxAddIn";
        private const string LogFileName = "startup.log";

        public static void Log(string message)
        {
            try
            {
                string folder = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    FolderName);
                Directory.CreateDirectory(folder);
                File.AppendAllText(
                    Path.Combine(folder, LogFileName),
                    DateTimeOffset.Now.ToString("o") + "\t" + message + Environment.NewLine);
            }
            catch
            {
            }
        }

        public static void LogException(string context, Exception exception)
        {
            if (exception == null)
            {
                Log(context + " failed with an unknown error.");
                return;
            }

            Log(context + " failed: " + exception);
        }

        public static void ShowError(string title, string context, Exception exception)
        {
            string message = context + " failed.";
            if (exception != null && !string.IsNullOrWhiteSpace(exception.Message))
            {
                message += Environment.NewLine + Environment.NewLine + exception.Message;
            }

            MessageBox.Show(
                GetExcelOwnerWindow(),
                message,
                title,
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
        }

        private static IWin32Window GetExcelOwnerWindow()
        {
            try
            {
                if (Globals.ExcelCSIToolBoxAddin != null &&
                    Globals.ExcelCSIToolBoxAddin.Application != null)
                {
                    IntPtr hwnd = new IntPtr(Globals.ExcelCSIToolBoxAddin.Application.Hwnd);
                    if (hwnd != IntPtr.Zero)
                    {
                        return new NativeWindowWrapper(hwnd);
                    }
                }
            }
            catch
            {
            }

            return null;
        }

        private sealed class NativeWindowWrapper : IWin32Window
        {
            public NativeWindowWrapper(IntPtr handle)
            {
                Handle = handle;
            }

            public IntPtr Handle { get; private set; }
        }
    }
}
