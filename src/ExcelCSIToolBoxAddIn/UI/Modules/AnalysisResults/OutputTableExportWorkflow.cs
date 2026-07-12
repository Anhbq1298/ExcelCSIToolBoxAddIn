using System;
using System.Diagnostics;
using System.IO;
using System.Windows;
using System.Windows.Interop;
using ExcelCSIToolBox.Application.Composition;
using ExcelCSIToolBox.Core.Abstractions.CSI;
using ExcelCSIToolBox.Core.Abstractions.Excel;
using ExcelCSIToolBoxAddIn.UI.Views;

namespace ExcelCSIToolBoxAddIn.UI.ViewModels
{
    public static class OutputTableExportWorkflow
    {
        public static OutputTableExportOptionsWindow Run(
            OutputTableExportConfig config,
            CsiToolboxUseCaseBundle useCases,
            ICSISapModelConnectionService csiConnectionService,
            IExcelOutputService excelOutputService,
            Window owner = null,
            IntPtr ownerHandle = default(IntPtr))
        {
            if (config == null)
            {
                throw new ArgumentNullException(nameof(config));
            }

            AnalysisExportDiagnostics.Log("Opening export options popup: " + config.TableDisplayName);
            var viewModel = new OutputTableExportOptionsViewModel(
                useCases,
                csiConnectionService,
                excelOutputService,
                config);
            var window = new OutputTableExportOptionsWindow(viewModel);
            if (owner != null && !ReferenceEquals(owner, window))
            {
                window.Owner = owner;
                window.WindowStartupLocation = WindowStartupLocation.CenterOwner;
            }
            else if (ownerHandle != IntPtr.Zero)
            {
                new WindowInteropHelper(window).Owner = ownerHandle;
                window.WindowStartupLocation = WindowStartupLocation.CenterOwner;
            }
            else
            {
                window.WindowStartupLocation = WindowStartupLocation.CenterScreen;
            }

            AnalysisExportDiagnostics.Log("Showing export options popup modally: " + config.TableDisplayName);
            bool? result = window.ShowDialog();
            AnalysisExportDiagnostics.Log(
                "Popup result: " + (result == true ? "Confirmed" : result == false ? "Cancelled" : "Closed") +
                " for " + config.TableDisplayName);
            return window;
        }
    }

    internal static class AnalysisExportDiagnostics
    {
        private static readonly object SyncRoot = new object();

        internal static void Log(string message)
        {
            string line = DateTimeOffset.Now.ToString("o") + "\t" + Safe(message);
            Trace.WriteLine("AnalysisExport: " + Safe(message));

            try
            {
                string root = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
                string folder = Path.Combine(root, "ExcelCSIToolBoxAddIn");
                Directory.CreateDirectory(folder);
                string path = Path.Combine(folder, "analysis-export.log");

                lock (SyncRoot)
                {
                    File.AppendAllText(path, line + Environment.NewLine);
                }
            }
            catch (Exception ex)
            {
                Trace.WriteLine("AnalysisExport logging failed: " + ex.Message);
            }
        }

        private static string Safe(string value)
        {
            return (value ?? string.Empty)
                .Replace("\r", " ")
                .Replace("\n", " ")
                .Replace("\t", " ");
        }
    }
}
