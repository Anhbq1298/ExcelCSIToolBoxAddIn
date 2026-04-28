using ExcelCSIToolBoxAddIn.AddIn;
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
            WindowManager.ShowEtabsWindow();
        }

        private void buttonSap2000_Click(object sender, RibbonControlEventArgs e)
        {
            WindowManager.ShowSap2000Window();
        }
    }
}
