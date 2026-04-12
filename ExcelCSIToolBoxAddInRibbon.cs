using Microsoft.Office.Tools.Ribbon;

namespace ExcelCSIToolBoxAddIn
{
    public partial class ExcelCSIToolBoxAddInRibbon
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
        }

        private void buttonEtabs_Click(object sender, RibbonControlEventArgs e)
        {
            // Ribbon click only delegates to launcher; no ETABS or UI business logic here.
            Globals.ExcelCSIToolBoxAddin.EtabsWindowLauncher?.OpenWindow();
        }
    }
}
