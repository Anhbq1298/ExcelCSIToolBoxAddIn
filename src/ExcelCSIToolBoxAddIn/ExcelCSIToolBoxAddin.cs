using ExcelCSIToolBox.Core.Abstractions;
using ExcelCSIToolBox.Infrastructure.CSI.Common.Adapters;
using ExcelCSIToolBox.Infrastructure.CSI.Etabs.Session;
using ExcelCSIToolBox.Infrastructure.CSI.Sap2000.Session;
using ExcelCSIToolBoxAddIn.AddIn;

namespace ExcelCSIToolBoxAddIn
{
    public partial class ExcelCSIToolBoxAddin
    {
        private void ExcelCSIToolBoxAddin_Startup(object sender, System.EventArgs e)
        {
            IThreadDispatcher threadDispatcher = new WpfThreadDispatcher();
            IProgressReporter progressReporter = new BatchProgressReporter(threadDispatcher);
            var etabsConnectionService = new EtabsConnectionService(new EtabsModelAdapter(), progressReporter);
            var sap2000ConnectionService = new Sap2000ConnectionService(new Sap2000ModelAdapter(), progressReporter);

            AddInCompositionRoot.Configure(etabsConnectionService, sap2000ConnectionService, progressReporter, threadDispatcher);
        }

        private void ExcelCSIToolBoxAddin_Shutdown(object sender, System.EventArgs e)
        {
            WindowManager.DisposePanes();
            AiTaskPaneManager.DisposePane();
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
