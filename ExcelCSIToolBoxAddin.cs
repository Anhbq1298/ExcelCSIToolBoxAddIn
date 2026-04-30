using ExcelCSIToolBox.Infrastructure.CSISapModel;
using ExcelCSIToolBox.Infrastructure.CSISapModel.Adapters;
using ExcelCSIToolBox.Infrastructure.Etabs;
using ExcelCSIToolBox.Infrastructure.Sap2000;
using ExcelCSIToolBoxAddIn.AddIn;
using ExcelCSIToolBoxAddIn.UI.Views;

namespace ExcelCSIToolBoxAddIn
{
    public partial class ExcelCSIToolBoxAddin
    {
        private void ExcelCSIToolBoxAddin_Startup(object sender, System.EventArgs e)
        {
            var etabsConnectionService = new EtabsConnectionService(new EtabsModelAdapter());
            var sap2000ConnectionService = new Sap2000ConnectionService(new Sap2000ModelAdapter());

            BatchProgressHost.ProgressRunner = BatchProgressWindow.RunForInfrastructure;
            AddInCompositionRoot.Configure(etabsConnectionService, sap2000ConnectionService);
        }

        private void ExcelCSIToolBoxAddin_Shutdown(object sender, System.EventArgs e)
        {
            BatchProgressHost.ProgressRunner = null;
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
