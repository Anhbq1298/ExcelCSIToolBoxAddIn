using System;
using System.Windows;
using ExcelCSIToolBox.Application.UseCases;
using ExcelCSIToolBox.Core.Models.AnalysisResults;
using ExcelCSIToolBox.Infrastructure.Services.Etabs.AnalysisResults;

namespace ExcelCSIToolBoxAddIn.UI.ViewModels
{
    public partial class CsiToolboxViewModel
    {
        private void OpenGetBaseReactionsDialog()
        {
            if (!PrepareExportWithGlobalUnit())
            {
                return;
            }

            OutputTableExportWorkflow.Run(
                CreateOutputTableExportConfig("Base Reactions"),
                _useCases,
                _csiConnectionService,
                _excelOutputService);
        }

        private void SelectWorkspacePage(string pageIndex)
        {
            if (pageIndex != null && pageIndex.StartsWith("Results:", StringComparison.OrdinalIgnoreCase))
            {
                SetTableGroup("ANALYSIS RESULTS", pageIndex.Substring("Results:".Length));
                ActiveWorkspacePage = 6;
                return;
            }

            if (pageIndex != null && pageIndex.StartsWith("Tables:", StringComparison.OrdinalIgnoreCase))
            {
                string value = pageIndex.Substring("Tables:".Length);
                string[] parts = value.Split(new[] { ':' }, 2);
                if (parts.Length == 2)
                {
                    SetTableGroup(parts[0], parts[1]);
                }
                else
                {
                    SetTableGroup("ANALYSIS RESULTS", value);
                }

                ActiveWorkspacePage = 6;
                return;
            }

            int index;
            if (int.TryParse(pageIndex, out index) && index >= 0 && index <= 8)
            {
                ActiveWorkspacePage = index;
                if (index == 8)
                {
                    return;
                }

                if (index == 7 && IsConnected)
                {
                    if (FrameStiffnessSections.Count == 0)
                    {
                        RefreshFrameStiffnessSections();
                    }

                    if (AreaStiffnessSections.Count == 0)
                    {
                        RefreshAreaStiffnessSections();
                    }
                }
            }
        }

        private async void RunAnalysisResult(AnalysisResultItem item)
        {
            if (item == null)
            {
                return;
            }

            try
            {
                await _analysisResultRouter.ExecuteAsync(item);
                StatusText = "Exported " + item.Title + " to Excel.";
            }
            catch (Exception ex)
            {
                string message = string.IsNullOrWhiteSpace(ex.Message)
                    ? "Failed to export ETABS analysis result."
                    : ex.Message;
                StatusText = message;
                MessageBox.Show(message, ProductTitle, MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private OutputTableExportConfig CreateOutputTableExportConfig(string displayTableName)
        {
            string tableName = string.IsNullOrWhiteSpace(displayTableName) ? "Base Reactions" : displayTableName;
            string groupName = string.Equals(tableName, "Base Reactions", StringComparison.OrdinalIgnoreCase)
                ? "Base Reactions"
                : string.IsNullOrWhiteSpace(ActiveAnalysisResultsGroup)
                ? tableName
                : ActiveAnalysisResultsGroup;
            string breadcrumb = string.Equals(groupName, tableName, StringComparison.OrdinalIgnoreCase)
                ? "ETABS Toolbox / ANALYSIS RESULTS / " + tableName
                : "ETABS Toolbox / ANALYSIS RESULTS / " + groupName + " / " + tableName;

            string popupProfileKey = "ForceOutput";
            if (string.Equals(groupName, "Objects and Elements", StringComparison.OrdinalIgnoreCase))
            {
                popupProfileKey = "ObjectsAndElements";
            }

            return new OutputTableExportConfig
            {
                TableDisplayName = tableName,
                Breadcrumb = breadcrumb,
                Description = "Select output cases to export " + tableName + ".",
                PopupProfileKey = popupProfileKey,
                ExportUnitOption = CreateExportUnitOption()
            };
        }

        private BaseReactionUnitOption CreateExportUnitOption()
        {
            return SelectedUnitSystem == null ? null : SelectedUnitSystem.ToExportUnitOption();
        }

        private void SetTableGroup(string category, string groupName)
        {
            ActiveTableCategory = string.IsNullOrWhiteSpace(category)
                ? "ANALYSIS RESULTS"
                : category;

            ActiveAnalysisResultsGroup = string.IsNullOrWhiteSpace(groupName)
                ? "Base Reactions"
                : groupName;

            AnalysisResultGroup group = EtabsAnalysisResultRegistry.CreateGroupForNavigation(ActiveAnalysisResultsGroup);
            ActiveAnalysisResultsGroup = group.Name;
            AnalysisResultTables.Clear();
            foreach (AnalysisResultItem item in group.Items)
            {
                AnalysisResultTables.Add(item);
            }

            AnalysisResultItem matchingItem = null;
            foreach (AnalysisResultItem item in AnalysisResultTables)
            {
                if (string.Equals(item.Title, groupName, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(item.EtabsTableName, groupName, StringComparison.OrdinalIgnoreCase))
                {
                    matchingItem = item;
                    break;
                }
            }

            SelectedAnalysisResultTable = matchingItem == null
                ? (AnalysisResultTables.Count > 0 ? AnalysisResultTables[0].Title : null)
                : matchingItem.Title;
        }

        private void OpenModalMassParticipationRatiosDialog()
        {
            if (!PrepareExportWithGlobalUnit())
            {
                return;
            }

            var viewModel = new GetModalMassParticipationRatiosViewModel(_useCases, _csiConnectionService, _excelOutputService);
            new ExcelCSIToolBoxAddIn.UI.Views.GetModalMassParticipationRatiosWindow(viewModel).Show();
        }

        private void OpenStoryForcesDialog()
        {
            OpenStoryResultsDialog(StoryPostprocessingResultKind.StoryForces);
        }

        private void OpenStoryDriftsDialog()
        {
            OpenStoryResultsDialog(StoryPostprocessingResultKind.StoryDrifts);
        }

        private void OpenStoryMaxOverAverageDisplacementsDialog()
        {
            OpenStoryResultsDialog(StoryPostprocessingResultKind.StoryMaxOverAverageDisplacements);
        }

        private void OpenStoryMaxOverAverageDriftsDialog()
        {
            OpenStoryResultsDialog(StoryPostprocessingResultKind.StoryMaxOverAverageDrifts);
        }

        private void OpenStoryResultsDialog(StoryPostprocessingResultKind kind)
        {
            if (!PrepareExportWithGlobalUnit())
            {
                return;
            }

            var viewModel = new GetStoryResultsViewModel(kind, _useCases, _csiConnectionService, _excelOutputService, CreateExportUnitOption());
            new ExcelCSIToolBoxAddIn.UI.Views.GetStoryResultsWindow(viewModel).Show();
        }

        private void OpenMassSummaryByStoryDialog()
        {
            if (!PrepareExportWithGlobalUnit())
            {
                return;
            }

            var viewModel = new GetMassSummaryByStoryViewModel(_useCases, _csiConnectionService, _excelOutputService);
            new ExcelCSIToolBoxAddIn.UI.Views.GetMassSummaryByStoryWindow(viewModel).Show();
        }
    }
}
