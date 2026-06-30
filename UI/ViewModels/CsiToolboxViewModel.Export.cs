using System;
using ExcelCSIToolBox.Application.UseCases;

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
            if (int.TryParse(pageIndex, out index) && index >= 0 && index <= 7)
            {
                ActiveWorkspacePage = index;
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

        private void ShowOutputSelectionAndExport(string displayTableName)
        {
            if (string.IsNullOrWhiteSpace(displayTableName))
            {
                return;
            }

            if (!PrepareExportWithGlobalUnit())
            {
                return;
            }

            if (string.Equals(ActiveAnalysisResultsGroup, "Modal Information", StringComparison.OrdinalIgnoreCase))
            {
                RunModalTableExport(displayTableName);
                return;
            }

            if (string.Equals(ActiveTableCategory, "MISCELLANEOUS DATA", StringComparison.OrdinalIgnoreCase))
            {
                RunMiscellaneousDataExport(displayTableName);
                return;
            }

            if (string.Equals(ActiveAnalysisResultsGroup, "Etabs Object Connectivity", StringComparison.OrdinalIgnoreCase))
            {
                RunEtabsObjectConnectivityExport(displayTableName);
                return;
            }

            if (string.Equals(ActiveAnalysisResultsGroup, "Other Output Items", StringComparison.OrdinalIgnoreCase))
            {
                RunOtherOutputItemsExport(displayTableName);
                return;
            }

            if (string.Equals(ActiveAnalysisResultsGroup, "Mass Data", StringComparison.OrdinalIgnoreCase))
            {
                RunMassDataExport(displayTableName);
                return;
            }

            if (string.Equals(ActiveAnalysisResultsGroup, "Displacements", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(ActiveAnalysisResultsGroup, "Reactions", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(ActiveAnalysisResultsGroup, "Velocity and Acceleration", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(ActiveAnalysisResultsGroup, "Joint Output", StringComparison.OrdinalIgnoreCase) ||
                ActiveAnalysisResultsGroup.StartsWith("Joint ", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(ActiveAnalysisResultsGroup, "Assembled Joint Masses", StringComparison.OrdinalIgnoreCase))
            {
                RunJointOutputExport(displayTableName);
                return;
            }

            OutputTableExportWorkflow.Run(
                CreateOutputTableExportConfig(displayTableName),
                _useCases,
                _csiConnectionService,
                _excelOutputService);
        }

        private void RunEtabsObjectConnectivityExport(string tableDisplayName)
        {
            OutputTableExportWorkflow.Run(
                new OutputTableExportConfig
                {
                    TableDisplayName = tableDisplayName,
                    Breadcrumb = "ETABS Toolbox / Element Manipulation / Etabs Object Connectivity / " + tableDisplayName,
                    Description = "Export ETABS " + tableDisplayName + ".",
                    PopupProfileKey = "EtabsObjectConnectivity",
                    ExportUnitOption = CreateExportUnitOption()
                },
                _useCases,
                _csiConnectionService,
                _excelOutputService);
        }

        private void RunModalTableExport(string tableDisplayName)
        {
            bool isResponseSpectrumModalInfo = string.Equals(
                tableDisplayName,
                "Response Spectrum Modal Info",
                StringComparison.OrdinalIgnoreCase);
            OutputTableExportWorkflow.Run(
                new OutputTableExportConfig
                {
                    TableDisplayName = tableDisplayName,
                    Breadcrumb = "ETABS Toolbox / ANALYSIS RESULTS / Modal Information / " + tableDisplayName,
                    Description = isResponseSpectrumModalInfo
                        ? "Select response spectrum case to export " + tableDisplayName + "."
                        : "Select modal case to export " + tableDisplayName + ".",
                    PopupProfileKey = isResponseSpectrumModalInfo
                        ? "ResponseSpectrumModalInfo"
                        : "ModalInformation",
                    ExportUnitOption = CreateExportUnitOption()
                },
                _useCases,
                _csiConnectionService,
                _excelOutputService);
        }

        private void RunMiscellaneousDataExport(string tableDisplayName)
        {
            bool isProjectInformation = string.Equals(
                tableDisplayName,
                "Project Information",
                StringComparison.OrdinalIgnoreCase);
            string groupName = isProjectInformation ? "Project Information" : "Material List";

            OutputTableExportWorkflow.Run(
                new OutputTableExportConfig
                {
                    TableDisplayName = tableDisplayName,
                    Breadcrumb = "ETABS Toolbox / MISCELLANEOUS DATA / " + groupName + " / " + tableDisplayName,
                    Description = isProjectInformation
                        ? "Export ETABS project information."
                        : "Export " + tableDisplayName + " using the main window unit system.",
                    PopupProfileKey = isProjectInformation ? "ProjectInformation" : "MaterialList",
                    ExportUnitOption = CreateExportUnitOption()
                },
                _useCases,
                _csiConnectionService,
                _excelOutputService);
        }

        private void RunOtherOutputItemsExport(string tableDisplayName)
        {
            if (string.Equals(tableDisplayName, "Story Forces", StringComparison.OrdinalIgnoreCase))
            {
                OutputTableExportWorkflow.Run(
                    new OutputTableExportConfig
                    {
                        TableDisplayName = "Story Forces",
                        Breadcrumb = "ETABS Toolbox / ANALYSIS RESULTS / Other Output Items / Story Forces",
                        Description = "Select load case or load combination to export Story Forces.",
                        PopupProfileKey = "StoryForces",
                        ExportUnitOption = CreateExportUnitOption()
                    },
                    _useCases,
                    _csiConnectionService,
                    _excelOutputService);
            }
            else if (string.Equals(tableDisplayName, "Diaphragm Forces", StringComparison.OrdinalIgnoreCase))
            {
                OutputTableExportWorkflow.Run(
                    new OutputTableExportConfig
                    {
                        TableDisplayName = "Diaphragm Forces",
                        Breadcrumb = "ETABS Toolbox / ANALYSIS RESULTS / Other Output Items / Diaphragm Forces",
                        Description = "Select load case or load combination to export Diaphragm Forces.",
                        PopupProfileKey = "DiaphragmForces",
                        ExportUnitOption = CreateExportUnitOption()
                    },
                    _useCases,
                    _csiConnectionService,
                    _excelOutputService);
            }
            else if (string.Equals(tableDisplayName, "Story Stiffness", StringComparison.OrdinalIgnoreCase))
            {
                OutputTableExportWorkflow.Run(
                    new OutputTableExportConfig
                    {
                        TableDisplayName = "Story Stiffness",
                        Breadcrumb = "ETABS Toolbox / ANALYSIS RESULTS / Other Output Items / Story Stiffness",
                        Description = "Select seismic, response spectrum, or wind load case to export Story Stiffness.",
                        PopupProfileKey = "SeismicWindOrRSOnlyWithUnit",
                        ExportUnitOption = CreateExportUnitOption()
                    },
                    _useCases,
                    _csiConnectionService,
                    _excelOutputService);
            }
            else if (string.Equals(tableDisplayName, "Shear Gravity Ratios", StringComparison.OrdinalIgnoreCase) ||
                     string.Equals(tableDisplayName, "Stiffness Gravity Ratios", StringComparison.OrdinalIgnoreCase))
            {
                OutputTableExportWorkflow.Run(
                    new OutputTableExportConfig
                    {
                        TableDisplayName = tableDisplayName,
                        Breadcrumb = "ETABS Toolbox / ANALYSIS RESULTS / Other Output Items / " + tableDisplayName,
                        Description = "Select seismic, response spectrum, or wind load case to export " + tableDisplayName + ".",
                        PopupProfileKey = "SeismicWindOrRSOnlyRatio",
                        ExportUnitOption = CreateExportUnitOption()
                    },
                    _useCases,
                    _csiConnectionService,
                    _excelOutputService);
            }
            else
            {
                OutputTableExportWorkflow.Run(
                    new OutputTableExportConfig
                    {
                        TableDisplayName = tableDisplayName,
                        Breadcrumb = "ETABS Toolbox / ANALYSIS RESULTS / Other Output Items / " + tableDisplayName,
                        Description = "Export " + tableDisplayName + " using the main window unit system.",
                        PopupProfileKey = "OtherOutputWithUnit",
                        ExportUnitOption = CreateExportUnitOption()
                    },
                    _useCases,
                    _csiConnectionService,
                    _excelOutputService);
            }
        }

        private void RunMassDataExport(string tableDisplayName)
        {
            OutputTableExportWorkflow.Run(
                new OutputTableExportConfig
                {
                    TableDisplayName = tableDisplayName,
                    Breadcrumb = "ETABS Toolbox / ANALYSIS RESULTS / Structure Output / Mass Data / " + tableDisplayName,
                    Description = "Export " + tableDisplayName + " using the main window unit system.",
                    PopupProfileKey = "MassData",
                    ExportUnitOption = CreateExportUnitOption()
                },
                _useCases,
                _csiConnectionService,
                _excelOutputService);
        }

        private void RunJointOutputExport(string tableDisplayName)
        {
            bool isJointMasses = string.Equals(tableDisplayName, "Assembled Joint Masses", StringComparison.OrdinalIgnoreCase);
            
            OutputTableExportWorkflow.Run(
                new OutputTableExportConfig
                {
                    TableDisplayName = tableDisplayName,
                    Breadcrumb = "ETABS Toolbox / ANALYSIS RESULTS / Joint Output / " + tableDisplayName,
                    Description = isJointMasses 
                        ? "Export Assembled Joint Masses using the main window unit system." 
                        : "Select load cases/combinations to export " + tableDisplayName + ".",
                    PopupProfileKey = isJointMasses ? "OtherOutputWithUnit" : "JointOutput",
                    ExportUnitOption = CreateExportUnitOption()
                },
                _useCases,
                _csiConnectionService,
                _excelOutputService);
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

            AnalysisResultTables.Clear();

            switch (ActiveAnalysisResultsGroup)
            {
                case "Base Reactions":
                    AnalysisResultTables.Add("Base Reactions");
                    break;
                case "Modal Information":
                    AnalysisResultTables.Add("Modal Periods And Frequencies");
                    AnalysisResultTables.Add("Modal Participating Mass Ratios");
                    AnalysisResultTables.Add("Modal Load Participation Ratios");
                    AnalysisResultTables.Add("Modal Participation Factors");
                    AnalysisResultTables.Add("Modal Direction Factors");
                    AnalysisResultTables.Add("Response Spectrum Modal Info");
                    break;
                case "Other Output Items":
                    AnalysisResultTables.Add("Centers Of Mass And Rigidity");
                    AnalysisResultTables.Add("Story Forces");
                    AnalysisResultTables.Add("Diaphragm Forces");
                    AnalysisResultTables.Add("Story Stiffness");
                    AnalysisResultTables.Add("Shear Gravity Ratios");
                    AnalysisResultTables.Add("Stiffness Gravity Ratios");
                    AnalysisResultTables.Add("Tributary Area and LLRF");
                    break;
                case "Mass Data":
                case "Mass Summary by Story":
                case "Mass Summary by Diaphragm":
                case "Mass Summary by Group":
                    ActiveAnalysisResultsGroup = "Mass Data";
                    AnalysisResultTables.Add("Mass Summary by Story");
                    AnalysisResultTables.Add("Mass Summary by Diaphragm");
                    AnalysisResultTables.Add("Mass Summary by Group");
                    break;
                case "Joint Output":
                case "Displacements":
                case "Joint Displacements":
                case "Joint Displacements - Absolute":
                case "Joint Drifts":
                case "Diaphragm Center Of Mass Displacements":
                case "Diaphragm Max Over Avg Drifts":
                case "Story Drifts":
                case "Story Max Over Avg Displacements":
                case "Story Max Over Avg Drifts":
                    ActiveAnalysisResultsGroup = "Displacements";
                    AnalysisResultTables.Add("Joint Displacements");
                    AnalysisResultTables.Add("Joint Displacements - Absolute");
                    AnalysisResultTables.Add("Joint Drifts");
                    AnalysisResultTables.Add("Diaphragm Center Of Mass Displacements");
                    AnalysisResultTables.Add("Diaphragm Max Over Avg Drifts");
                    AnalysisResultTables.Add("Story Drifts");
                    AnalysisResultTables.Add("Story Max Over Avg Displacements");
                    AnalysisResultTables.Add("Story Max Over Avg Drifts");
                    break;
                case "Reactions":
                case "Joint Reactions":
                case "Joint Design Reactions":
                    ActiveAnalysisResultsGroup = "Reactions";
                    AnalysisResultTables.Add("Joint Reactions");
                    AnalysisResultTables.Add("Joint Design Reactions");
                    break;
                case "Velocity and Acceleration":
                case "Joint Velocities - Relative":
                case "Joint Velocities - Absolute":
                case "Joint Accelerations - Relative":
                case "Joint Accelerations - Absolute":
                case "Diaphragm Accelerations":
                case "Story Accelerations":
                    ActiveAnalysisResultsGroup = "Velocity and Acceleration";
                    AnalysisResultTables.Add("Joint Velocities - Relative");
                    AnalysisResultTables.Add("Joint Velocities - Absolute");
                    AnalysisResultTables.Add("Joint Accelerations - Relative");
                    AnalysisResultTables.Add("Joint Accelerations - Absolute");
                    AnalysisResultTables.Add("Diaphragm Accelerations");
                    AnalysisResultTables.Add("Story Accelerations");
                    break;
                case "Frame Output":
                    AnalysisResultTables.Add("Element Forces - Columns");
                    AnalysisResultTables.Add("Element Forces - Beams");
                    AnalysisResultTables.Add("Element Forces - Braces");
                    AnalysisResultTables.Add("Element Joint Forces - Frame");
                    break;
                case "Area Output":
                    AnalysisResultTables.Add("Element Forces - Area Shells");
                    AnalysisResultTables.Add("Element Stresses - Area Shells");
                    AnalysisResultTables.Add("Element Strains - Area Shells");
                    AnalysisResultTables.Add("Element Joint Forces - Shells");
                    break;
                case "Wall Output":
                    AnalysisResultTables.Add("Pier Forces");
                    break;
                case "Objects and Elements":
                    AnalysisResultTables.Add("Objects and Elements - Joints");
                    AnalysisResultTables.Add("Objects and Elements - Frames");
                    AnalysisResultTables.Add("Objects and Elements - Areas");
                    break;
                case "Etabs Object Connectivity":
                case "Point Object Connectivity":
                case "Beam Object Connectivity":
                case "Column Object Connectivity":
                case "Brace Object Connectivity":
                case "Floor Object Connectivity":
                case "Wall Object Connectivity":
                    ActiveAnalysisResultsGroup = "Etabs Object Connectivity";
                    AnalysisResultTables.Add("Point Object Connectivity");
                    AnalysisResultTables.Add("Beam Object Connectivity");
                    AnalysisResultTables.Add("Column Object Connectivity");
                    AnalysisResultTables.Add("Brace Object Connectivity");
                    AnalysisResultTables.Add("Floor Object Connectivity");
                    AnalysisResultTables.Add("Wall Object Connectivity");
                    break;
                case "Assembled Joint Masses":
                    ActiveAnalysisResultsGroup = "Assembled Joint Masses";
                    AnalysisResultTables.Add("Assembled Joint Masses");
                    break;
                case "Project Information":
                    AnalysisResultTables.Add("Project Information");
                    break;
                case "Material List":
                    AnalysisResultTables.Add("Material List by Object Type");
                    AnalysisResultTables.Add("Material List by Section Property");
                    AnalysisResultTables.Add("Material List by Story");
                    break;
            }

            string matchingTable = null;
            if (AnalysisResultTables.Count > 0)
            {
                foreach (string table in AnalysisResultTables)
                {
                    if (string.Equals(table, groupName, StringComparison.OrdinalIgnoreCase))
                    {
                        matchingTable = table;
                        break;
                    }
                }
                if (matchingTable == null)
                {
                    if (string.Equals(groupName, "Joint Output", StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(groupName, "Displacements", StringComparison.OrdinalIgnoreCase))
                    {
                        matchingTable = "Joint Displacements";
                    }
                    else if (string.Equals(groupName, "Reactions", StringComparison.OrdinalIgnoreCase))
                    {
                        matchingTable = "Joint Reactions";
                    }
                    else if (string.Equals(groupName, "Velocity and Acceleration", StringComparison.OrdinalIgnoreCase))
                    {
                        matchingTable = "Joint Velocities - Relative";
                    }
                }
            }

            SelectedAnalysisResultTable = matchingTable ?? (AnalysisResultTables.Count > 0 ? AnalysisResultTables[0] : null);
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
