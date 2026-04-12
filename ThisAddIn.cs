using ExcelCSIToolBoxAddIn.AddIn;
using ExcelCSIToolBoxAddIn.Infrastructure.Etabs;

namespace ExcelCSIToolBoxAddIn
{
    public partial class ThisAddIn
    {
        internal EtabsToolboxWindowLauncher EtabsWindowLauncher { get; private set; }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            // Lightweight composition root for phase 1.
            var etabsConnectionService = new EtabsConnectionService();
            EtabsWindowLauncher = new EtabsToolboxWindowLauncher(etabsConnectionService);
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
