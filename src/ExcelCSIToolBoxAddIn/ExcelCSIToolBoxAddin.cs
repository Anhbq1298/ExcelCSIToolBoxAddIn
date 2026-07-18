using ExcelCSIToolBox.Core.Abstractions;
using ExcelCSIToolBox.Infrastructure.CSI.Etabs.Session;
using ExcelCSIToolBox.Infrastructure.CSI.Sap2000.Session;
using ExcelCSIToolBoxAddIn.AddIn;
using System;
using System.IO;
using System.Windows.Forms;

namespace ExcelCSIToolBoxAddIn
{
    public partial class ExcelCSIToolBoxAddin
    {
        private static Timer _startupPaneTimer;
        private static int _startupPaneAttempts;

        private void ExcelCSIToolBoxAddin_Startup(object sender, System.EventArgs e)
        {
            LogStartup("Startup begin.");
            IThreadDispatcher threadDispatcher = new WpfThreadDispatcher();
            IProgressReporter progressReporter = new BatchProgressReporter(threadDispatcher);
            var etabsConnectionService = new EtabsConnectionService(new EtabsModelAdapter(), progressReporter);
            var sap2000ConnectionService = new Sap2000ConnectionService(new Sap2000ModelAdapter(), progressReporter);

            AddInCompositionRoot.Configure(etabsConnectionService, sap2000ConnectionService, progressReporter, threadDispatcher);
            LogStartup("Composition configured.");
            ShowEtabsToolboxAfterStartup();
        }

        private void ExcelCSIToolBoxAddin_Shutdown(object sender, System.EventArgs e)
        {
            ModelessWpfWindowService.CloseAll();
            WindowManager.DisposePanes();
            AiTaskPaneManager.DisposePane();
        }

        private static void ShowEtabsToolboxAfterStartup()
        {
            if (_startupPaneTimer != null)
            {
                _startupPaneTimer.Stop();
                _startupPaneTimer.Dispose();
            }

            _startupPaneTimer = new Timer();
            _startupPaneAttempts = 0;
            _startupPaneTimer.Interval = 1000;
            _startupPaneTimer.Tick += delegate
            {
                _startupPaneAttempts++;

                try
                {
                    if (Globals.ExcelCSIToolBoxAddin == null ||
                        Globals.ExcelCSIToolBoxAddin.Application == null ||
                        Globals.ExcelCSIToolBoxAddin.Application.Workbooks.Count == 0)
                    {
                        LogStartup("Waiting for an Excel workbook before opening ETABS Toolbox pane. Attempt " + _startupPaneAttempts + ".");
                        return;
                    }

                    _startupPaneTimer.Stop();
                    _startupPaneTimer.Dispose();
                    _startupPaneTimer = null;

                    LogStartup("Opening ETABS Toolbox pane.");
                    WindowManager.ShowEtabsWindow();
                    LogStartup("ETABS Toolbox pane command completed.");
                }
                catch (Exception ex)
                {
                    LogStartup("ETABS Toolbox pane command will retry after failure: " + ex.Message);
                    // Keep Excel startup alive if the task pane cannot be opened yet.
                }
            };
            _startupPaneTimer.Start();
        }

        private static void LogStartup(string message)
        {
            try
            {
                string folder = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "ExcelCSIToolBoxAddIn");
                Directory.CreateDirectory(folder);
                File.AppendAllText(
                    Path.Combine(folder, "startup.log"),
                    DateTimeOffset.Now.ToString("o") + "\t" + message + Environment.NewLine);
            }
            catch
            {
            }
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ExcelCSIToolBoxAddin_Startup);
            this.Shutdown += new System.EventHandler(ExcelCSIToolBoxAddin_Shutdown);
        }

        #endregion
    }
}
