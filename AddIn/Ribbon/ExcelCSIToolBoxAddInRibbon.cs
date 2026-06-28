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
                groupEtabsPostprocessing.Visible = false;

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

        private void buttonGetBaseReactions_Click(object sender, RibbonControlEventArgs e)
        {
            WindowManager.ShowGetBaseReactionsWindow();
        }

        private void buttonModalMassParticipationRatios_Click(object sender, RibbonControlEventArgs e)
        {
            WindowManager.ShowModalMassParticipationRatiosWindow();
        }

        private void buttonStoryForces_Click(object sender, RibbonControlEventArgs e)
        {
            WindowManager.ShowStoryForcesWindow();
        }

        private void buttonStoryDrifts_Click(object sender, RibbonControlEventArgs e)
        {
            WindowManager.ShowStoryDriftsWindow();
        }

        private void buttonStoryMaxOverAverageDisplacements_Click(object sender, RibbonControlEventArgs e)
        {
            WindowManager.ShowStoryMaxOverAverageDisplacementsWindow();
        }

        private void buttonStoryMaxOverAverageDrifts_Click(object sender, RibbonControlEventArgs e)
        {
            WindowManager.ShowStoryMaxOverAverageDriftsWindow();
        }

        private void buttonMassSummaryByStory_Click(object sender, RibbonControlEventArgs e)
        {
            WindowManager.ShowMassSummaryByStoryWindow();
        }

        private void buttonAiAgent_Click(object sender, RibbonControlEventArgs e)
        {
            AiTaskPaneManager.TogglePane();
        }
    }
}

