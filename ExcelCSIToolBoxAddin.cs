using ExcelCSIToolBoxAddIn.AddIn;
using ExcelCSIToolBoxAddIn.Infrastructure.Etabs;
using ExcelCSIToolBoxAddIn.Infrastructure.Excel;

namespace ExcelCSIToolBoxAddIn
{
    public partial class ExcelCSIToolBoxAddin
    {
        private void ExcelCSIToolBoxAddin_Startup(object sender, System.EventArgs e)
        {
            // Lightweight composition root for phase 1.
            var etabsConnectionService = new EtabsConnectionService();
            var sap2000ConnectionService = new Sap2000ConnectionService();
            var excelSelectionService = new ExcelSelectionService();
            var excelOutputService = new ExcelOutputService();
            WindowManager.Configure(etabsConnectionService, sap2000ConnectionService, excelSelectionService, excelOutputService);
        }

        private void ExcelCSIToolBoxAddin_Shutdown(object sender, System.EventArgs e)
        {
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
