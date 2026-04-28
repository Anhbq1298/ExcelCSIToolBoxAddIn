using ExcelCSIToolBoxAddIn.AddIn;
using Microsoft.Office.Tools.Ribbon;

namespace ExcelCSIToolBoxAddIn
{
    public partial class ExcelCSIToolBoxAddInRibbon
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            try
            {
                // In VSTO, the base directory is usually more reliable than Assembly.Location
                string baseDir = System.AppDomain.CurrentDomain.BaseDirectory;
                
                string etabsIconPath = System.IO.Path.Combine(baseDir, "icon", "etabs.png");
                if (System.IO.File.Exists(etabsIconPath))
                {
                    using (var stream = new System.IO.FileStream(etabsIconPath, System.IO.FileMode.Open, System.IO.FileAccess.Read))
                    {
                        this.buttonEtabs.Image = System.Drawing.Image.FromStream(stream);
                    }
                }

                string sapIconPath = System.IO.Path.Combine(baseDir, "icon", "sap2000icon.jpg");
                if (System.IO.File.Exists(sapIconPath))
                {
                    using (var stream = new System.IO.FileStream(sapIconPath, System.IO.FileMode.Open, System.IO.FileAccess.Read))
                    {
                        this.buttonSap2000.Image = System.Drawing.Image.FromStream(stream);
                    }
                }
            }
            catch
            {
                // Silently fail
            }
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

