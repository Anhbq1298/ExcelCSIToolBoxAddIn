using System;
using System.Windows;
using ExcelCSIToolBox.Application.UseCases;
using ExcelCSIToolBox.Core.Models.AnalysisResults;
using ExcelCSIToolBox.Core.Models.ElementConnectivity;
using ExcelCSIToolBox.Core.Models.MiscellaneousData;
using ExcelCSIToolBox.Infrastructure.Services.Etabs.AnalysisResults;
using ExcelCSIToolBox.Infrastructure.Services.Etabs.ElementConnectivity;
using ExcelCSIToolBox.Infrastructure.Services.Etabs.MiscellaneousData;

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
            if (int.TryParse(pageIndex, out index) && index >= 0 && index <= 9)
            {
                ActiveWorkspacePage = index;
                if (index == 8 || index == 9)
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

        private void RunAnalysisResult(AnalysisResultItem item)
        {
            if (item == null)
            {
                return;
            }

            string tableName = GetAnalysisResultTableName(item);
            AnalysisExportDiagnostics.Log("Analysis export item clicked: " + tableName);

            try
            {
                if (!PrepareExportWithGlobalUnit())
                {
                    AnalysisExportDiagnostics.Log("Analysis export cancelled before popup because ETABS was not ready: " + tableName);
                    return;
                }

                OutputTableExportConfig config = CreateAnalysisResultExportConfig(item);
                StatusText = "Opening export options for " + tableName + ".";
                AnalysisExportDiagnostics.Log("Command execution resolved report type: " + tableName);
                OutputTableExportWorkflow.Run(
                    config,
                    _useCases,
                    _csiConnectionService,
                    _excelOutputService,
                    GetActiveOwnerWindow());
            }
            catch (Exception ex)
            {
                string message = string.IsNullOrWhiteSpace(ex.Message)
                    ? "Failed to open the ETABS analysis export options."
                    : ex.Message;
                StatusText = message;
                AnalysisExportDiagnostics.Log("Failed to open export options for " + tableName + ": " + message);
                MessageBox.Show(message, ProductTitle, MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private async void RunMiscellaneousData(MiscellaneousDataItem item)
        {
            if (item == null)
            {
                return;
            }

            try
            {
                await _miscellaneousDataRouter.ExecuteAsync(item);
                StatusText = "Exported " + item.Title + " to Excel.";
            }
            catch (Exception ex)
            {
                string message = string.IsNullOrWhiteSpace(ex.Message)
                    ? "Failed to export ETABS miscellaneous data."
                    : ex.Message;
                StatusText = message;
                MessageBox.Show(message, ProductTitle, MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private async void RunElementConnectivity(ElementConnectivityItem item)
        {
            if (item == null)
            {
                return;
            }

            try
            {
                await _elementConnectivityRouter.ExecuteAsync(item);
                StatusText = "Exported " + item.Title + " to Excel.";
            }
            catch (Exception ex)
            {
                string message = string.IsNullOrWhiteSpace(ex.Message)
                    ? "Failed to export ETABS element connectivity."
                    : ex.Message;
                StatusText = message;
                MessageBox.Show(message, ProductTitle, MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private void RunEtabsTableItem(object item)
        {
            AnalysisResultItem analysisResultItem = item as AnalysisResultItem;
            if (analysisResultItem != null)
            {
                RunAnalysisResult(analysisResultItem);
                return;
            }

            ElementConnectivityItem elementConnectivityItem = item as ElementConnectivityItem;
            if (elementConnectivityItem != null)
            {
                RunElementConnectivity(elementConnectivityItem);
                return;
            }

            MiscellaneousDataItem miscellaneousDataItem = item as MiscellaneousDataItem;
            if (miscellaneousDataItem != null)
            {
                RunMiscellaneousData(miscellaneousDataItem);
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

        private OutputTableExportConfig CreateAnalysisResultExportConfig(AnalysisResultItem item)
        {
            string tableName = GetAnalysisResultTableName(item);
            string groupName = ResolveAnalysisResultExportGroup(item);

            if (string.Equals(tableName, "Base Reactions", StringComparison.OrdinalIgnoreCase))
            {
                return new OutputTableExportConfig
                {
                    TableDisplayName = tableName,
                    Breadcrumb = "ETABS Toolbox / ANALYSIS RESULTS / Base Reactions",
                    Description = "Select load cases/combinations and output unit to export " + tableName + ".",
                    PopupProfileKey = "ForceOutput",
                    ExportUnitOption = CreateExportUnitOption()
                };
            }

            if (string.Equals(groupName, "Modal Information", StringComparison.OrdinalIgnoreCase))
            {
                bool isResponseSpectrumModalInfo = string.Equals(
                    tableName,
                    "Response Spectrum Modal Info",
                    StringComparison.OrdinalIgnoreCase);

                return new OutputTableExportConfig
                {
                    TableDisplayName = tableName,
                    Breadcrumb = "ETABS Toolbox / ANALYSIS RESULTS / Modal Information / " + tableName,
                    Description = isResponseSpectrumModalInfo
                        ? "Select response spectrum case to export " + tableName + "."
                        : "Select modal case to export " + tableName + ".",
                    PopupProfileKey = isResponseSpectrumModalInfo
                        ? "ResponseSpectrumModalInfo"
                        : "ModalInformation",
                    ExportUnitOption = CreateExportUnitOption()
                };
            }

            if (string.Equals(groupName, "Other Output Items", StringComparison.OrdinalIgnoreCase))
            {
                return CreateOtherOutputItemsExportConfig(tableName);
            }

            if (string.Equals(groupName, "Mass Data", StringComparison.OrdinalIgnoreCase))
            {
                return new OutputTableExportConfig
                {
                    TableDisplayName = tableName,
                    Breadcrumb = "ETABS Toolbox / ANALYSIS RESULTS / Structure Output / Mass Data / " + tableName,
                    Description = "Select output unit to export " + tableName + ".",
                    PopupProfileKey = "MassData",
                    ExportUnitOption = CreateExportUnitOption()
                };
            }

            if (IsJointOutputGroup(groupName))
            {
                bool isJointMasses = string.Equals(tableName, "Assembled Joint Masses", StringComparison.OrdinalIgnoreCase);
                return new OutputTableExportConfig
                {
                    TableDisplayName = tableName,
                    Breadcrumb = "ETABS Toolbox / ANALYSIS RESULTS / Joint Output / " + groupName + " / " + tableName,
                    Description = isJointMasses
                        ? "Select output unit to export " + tableName + "."
                        : "Select load cases/combinations and output unit to export " + tableName + ".",
                    PopupProfileKey = isJointMasses ? "OtherOutputWithUnit" : "JointOutput",
                    ExportUnitOption = CreateExportUnitOption()
                };
            }

            return CreateOutputTableExportConfig(tableName);
        }

        private OutputTableExportConfig CreateOtherOutputItemsExportConfig(string tableName)
        {
            if (string.Equals(tableName, "Story Forces", StringComparison.OrdinalIgnoreCase))
            {
                return CreateNamedAnalysisOutputConfig(tableName, "StoryForces", "Select load cases/combinations and output unit to export ");
            }

            if (string.Equals(tableName, "Diaphragm Forces", StringComparison.OrdinalIgnoreCase))
            {
                return CreateNamedAnalysisOutputConfig(tableName, "DiaphragmForces", "Select load cases/combinations and output unit to export ");
            }

            if (string.Equals(tableName, "Story Stiffness", StringComparison.OrdinalIgnoreCase))
            {
                return CreateNamedAnalysisOutputConfig(tableName, "SeismicWindOrRSOnlyWithUnit", "Select seismic, wind, or response spectrum cases and output unit to export ");
            }

            if (string.Equals(tableName, "Shear Gravity Ratios", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(tableName, "Stiffness Gravity Ratios", StringComparison.OrdinalIgnoreCase))
            {
                return CreateNamedAnalysisOutputConfig(tableName, "SeismicWindOrRSOnlyRatio", "Select seismic, wind, or response spectrum cases to export ");
            }

            return new OutputTableExportConfig
            {
                TableDisplayName = tableName,
                Breadcrumb = "ETABS Toolbox / ANALYSIS RESULTS / Other Output Items / " + tableName,
                Description = "Select output unit to export " + tableName + ".",
                PopupProfileKey = "OtherOutputWithUnit",
                ExportUnitOption = CreateExportUnitOption()
            };
        }

        private OutputTableExportConfig CreateNamedAnalysisOutputConfig(
            string tableName,
            string popupProfileKey,
            string descriptionPrefix)
        {
            return new OutputTableExportConfig
            {
                TableDisplayName = tableName,
                Breadcrumb = "ETABS Toolbox / ANALYSIS RESULTS / Other Output Items / " + tableName,
                Description = descriptionPrefix + tableName + ".",
                PopupProfileKey = popupProfileKey,
                ExportUnitOption = CreateExportUnitOption()
            };
        }

        private string ResolveAnalysisResultExportGroup(AnalysisResultItem item)
        {
            if (item == null)
            {
                return ActiveAnalysisResultsGroup;
            }

            if (!string.IsNullOrWhiteSpace(item.Category))
            {
                if (string.Equals(item.Category, "Joint Masses", StringComparison.OrdinalIgnoreCase))
                {
                    return "Assembled Joint Masses";
                }

                return item.Category;
            }

            return string.IsNullOrWhiteSpace(ActiveAnalysisResultsGroup)
                ? string.Empty
                : ActiveAnalysisResultsGroup;
        }

        private static bool IsJointOutputGroup(string groupName)
        {
            return string.Equals(groupName, "Displacements", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(groupName, "Reactions", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(groupName, "Velocity and Acceleration", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(groupName, "Assembled Joint Masses", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(groupName, "Joint Output", StringComparison.OrdinalIgnoreCase);
        }

        private static string GetAnalysisResultTableName(AnalysisResultItem item)
        {
            if (item == null)
            {
                return string.Empty;
            }

            return string.IsNullOrWhiteSpace(item.EtabsTableName)
                ? item.Title
                : item.EtabsTableName;
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

            EtabsTableItems.Clear();
            AnalysisResultTables.Clear();

            if (string.Equals(ActiveTableCategory, "Element Manipulation", StringComparison.OrdinalIgnoreCase))
            {
                SetElementConnectivityGroup(groupName);
                return;
            }

            if (string.Equals(ActiveTableCategory, "MISCELLANEOUS DATA", StringComparison.OrdinalIgnoreCase))
            {
                SetMiscellaneousDataGroup(groupName);
                return;
            }

            AnalysisResultGroup group = EtabsAnalysisResultRegistry.CreateGroupForNavigation(ActiveAnalysisResultsGroup);
            ActiveAnalysisResultsGroup = group.Name;
            foreach (AnalysisResultItem item in group.Items)
            {
                AnalysisResultTables.Add(item);
                EtabsTableItems.Add(item);
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

        private void SetElementConnectivityGroup(string groupName)
        {
            ElementConnectivityGroup group = EtabsElementConnectivityRegistry.CreateGroupForNavigation(groupName);
            ActiveAnalysisResultsGroup = group.Name;
            foreach (ElementConnectivityItem item in group.Items)
            {
                EtabsTableItems.Add(item);
            }

            ElementConnectivityItem matchingItem = null;
            foreach (ElementConnectivityItem item in group.Items)
            {
                if (string.Equals(item.Title, groupName, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(item.EtabsTableName, groupName, StringComparison.OrdinalIgnoreCase))
                {
                    matchingItem = item;
                    break;
                }
            }

            SelectedAnalysisResultTable = matchingItem == null
                ? (group.Items.Count > 0 ? group.Items[0].Title : null)
                : matchingItem.Title;
        }

        private void SetMiscellaneousDataGroup(string groupName)
        {
            MiscellaneousDataGroup group = EtabsMiscellaneousDataRegistry.CreateGroupForNavigation(groupName);
            ActiveAnalysisResultsGroup = group.Name;
            foreach (MiscellaneousDataItem item in group.Items)
            {
                EtabsTableItems.Add(item);
            }

            MiscellaneousDataItem matchingItem = null;
            foreach (MiscellaneousDataItem item in group.Items)
            {
                if (string.Equals(item.Title, groupName, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(item.EtabsTableName, groupName, StringComparison.OrdinalIgnoreCase))
                {
                    matchingItem = item;
                    break;
                }
            }

            SelectedAnalysisResultTable = matchingItem == null
                ? (group.Items.Count > 0 ? group.Items[0].Title : null)
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
