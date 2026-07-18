using System;
using ExcelCSIToolBoxAddIn.AddIn;
using ExcelCSIToolBoxAddIn.AddIn.Diagnostics;
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
                buttonAiAgent.Visible = false;

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
            RunRibbonAction("ETABS Toolbox", WindowManager.ShowEtabsWindow);
        }

        private void buttonGetBaseReactions_Click(object sender, RibbonControlEventArgs e)
        {
            RunRibbonAction("Get Base Reactions", WindowManager.ShowGetBaseReactionsWindow);
        }

        private void buttonModalMassParticipationRatios_Click(object sender, RibbonControlEventArgs e)
        {
            RunRibbonAction("Modal Mass Participation Ratios", WindowManager.ShowModalMassParticipationRatiosWindow);
        }

        private void buttonStoryForces_Click(object sender, RibbonControlEventArgs e)
        {
            RunRibbonAction("Story Forces", WindowManager.ShowStoryForcesWindow);
        }

        private void buttonStoryDrifts_Click(object sender, RibbonControlEventArgs e)
        {
            RunRibbonAction("Story Drifts", WindowManager.ShowStoryDriftsWindow);
        }

        private void buttonStoryMaxOverAverageDisplacements_Click(object sender, RibbonControlEventArgs e)
        {
            RunRibbonAction("Story Max/Avg Displacements", WindowManager.ShowStoryMaxOverAverageDisplacementsWindow);
        }

        private void buttonStoryMaxOverAverageDrifts_Click(object sender, RibbonControlEventArgs e)
        {
            RunRibbonAction("Story Max/Avg Drifts", WindowManager.ShowStoryMaxOverAverageDriftsWindow);
        }

        private void buttonMassSummaryByStory_Click(object sender, RibbonControlEventArgs e)
        {
            RunRibbonAction("Mass Summary by Story", WindowManager.ShowMassSummaryByStoryWindow);
        }

        private void buttonAiAgent_Click(object sender, RibbonControlEventArgs e)
        {
            RunRibbonAction("MHT AI Assistant", AiTaskPaneManager.TogglePane);
        }

        private void buttonRefreshPlugin_Click(object sender, RibbonControlEventArgs e)
        {
            RunRibbonAction("Refresh Plugin", AddInCompositionRoot.RefreshPlugin);
        }

        private void buttonAbout_Click(object sender, RibbonControlEventArgs e)
        {
            RunRibbonAction("About", WindowManager.ShowAboutWindow);
        }

        private void buttonDropPanel_Click(object sender, RibbonControlEventArgs e)
        {
            RunRibbonAction("Drop Panel", WindowManager.ShowDropPanelWindow);
        }

        private void buttonSap2000_Click(object sender, RibbonControlEventArgs e)
        {
            RunRibbonAction("SAP2000 Toolbox", WindowManager.ShowSap2000Window);
        }

        private static void RunRibbonAction(string actionName, Action action)
        {
            AddInDiagnostics.Log("Ribbon click received: " + actionName + ".");

            try
            {
                action();
                AddInDiagnostics.Log("Ribbon action completed: " + actionName + ".");
            }
            catch (Exception ex)
            {
                AddInDiagnostics.LogException("Ribbon action " + actionName, ex);
                AddInDiagnostics.ShowError("Excel CSI ToolBox", actionName, ex);
            }
        }
    }
}

