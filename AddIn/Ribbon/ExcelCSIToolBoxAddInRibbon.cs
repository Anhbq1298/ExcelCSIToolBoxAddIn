using ExcelCSIToolBoxAddIn.AddIn;
using ExcelCSIToolBoxAddIn.AddIn.Ribbon;
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

                var baseReactionsImg = LoadIcon(baseDir, "GetBaseReactions.ico");
                buttonGetBaseReactions.Image = baseReactionsImg ?? PostprocessingRibbonIcons.BaseReactions;

                var modalMassImg = LoadIcon(baseDir, "ModalMassParticipationRatios.ico");
                buttonModalMassParticipationRatios.Image = modalMassImg ?? PostprocessingRibbonIcons.ModalMassParticipation;

                var storyForcesImg = LoadIcon(baseDir, "StoryForces.ico");
                buttonStoryForces.Image = storyForcesImg ?? PostprocessingRibbonIcons.StoryForces;

                var storyDisplacementsImg = LoadIcon(baseDir, "StoryDisplacements.ico");
                if (storyDisplacementsImg != null)
                {
                    buttonStoryDrifts.Image = storyDisplacementsImg;
                    buttonStoryMaxOverAverageDisplacements.Image = storyDisplacementsImg;
                    buttonStoryMaxOverAverageDrifts.Image = storyDisplacementsImg;
                }
                else
                {
                    buttonStoryDrifts.Image = PostprocessingRibbonIcons.StoryDrifts;
                    buttonStoryMaxOverAverageDisplacements.Image = PostprocessingRibbonIcons.StoryMaxOverAverageDisplacements;
                    buttonStoryMaxOverAverageDrifts.Image = PostprocessingRibbonIcons.StoryMaxOverAverageDrifts;
                }
            }
            catch
            {
                // Silently fail
            }
        }

        private System.Drawing.Image LoadIcon(string baseDir, string filename)
        {
            string path = System.IO.Path.Combine(baseDir, "icon", filename);
            if (System.IO.File.Exists(path))
            {
                try
                {
                    using (var icon = new System.Drawing.Icon(path, 32, 32))
                    {
                        return icon.ToBitmap();
                    }
                }
                catch
                {
                    try
                    {
                        return System.Drawing.Image.FromFile(path);
                    }
                    catch
                    {
                        return null;
                    }
                }
            }
            return null;
        }

        private void buttonEtabs_Click(object sender, RibbonControlEventArgs e)
        {
            WindowManager.ShowEtabsWindow();
        }

        private void buttonSap2000_Click(object sender, RibbonControlEventArgs e)
        {
            WindowManager.ShowSap2000Window();
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

