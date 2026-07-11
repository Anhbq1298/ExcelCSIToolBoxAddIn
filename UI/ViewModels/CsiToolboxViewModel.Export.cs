using System;
using System.Collections.Generic;
using System.Globalization;
using System.Windows;
using ExcelCSIToolBox.Application.UseCases;
using ExcelCSIToolBox.Core.Abstractions.CSI;
using ExcelCSIToolBox.Core.Common.Results;
using ExcelCSIToolBox.Core.Models.AnalysisResults;
using ExcelCSIToolBox.Core.Models.ElementConnectivity;
using ExcelCSIToolBox.Core.Models.EtabsTables;
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
                if (!PrepareExportWithGlobalUnit())
                {
                    return;
                }

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
                if (!PrepareExportWithGlobalUnit())
                {
                    return;
                }

                OperationResult<IReadOnlyList<CsiSelectedObjectDto>> selectedResult =
                    _csiConnectionService.GetSelectedObjectsFromActiveModel();
                if (!selectedResult.IsSuccess)
                {
                    string selectionMessage = string.IsNullOrWhiteSpace(selectedResult.Message)
                        ? "Failed to read the current ETABS selection."
                        : selectedResult.Message;
                    StatusText = selectionMessage;
                    MessageBox.Show(selectionMessage, ProductTitle, MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                HashSet<string> selectedObjectNames = CreateSelectedObjectNameSet(selectedResult.Data);
                if (selectedObjectNames.Count == 0)
                {
                    string selectionMessage = "Select one or more ETABS objects before exporting " + item.Title + ".";
                    StatusText = selectionMessage;
                    MessageBox.Show(selectionMessage, ProductTitle, MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                StatusText = "Reading " + item.Title + " for the current ETABS selection.";
                var etabsConnectionService =
                    new ExcelCSIToolBox.Infrastructure.Services.Etabs.EtabsConnectionService(_csiConnectionService);
                var tableService =
                    new ExcelCSIToolBox.Infrastructure.Services.Etabs.EtabsDatabaseTableService(etabsConnectionService);
                EtabsTableResult fullTable = await tableService.GetTableAsync(item.EtabsTableName);
                EtabsTableResult selectedTable = FilterConnectivityRowsBySelection(fullTable, selectedObjectNames);
                if (selectedTable.Rows.Count == 0)
                {
                    string emptyMessage = "No " + item.Title + " rows matched the current ETABS selection.";
                    StatusText = emptyMessage;
                    MessageBox.Show(emptyMessage, ProductTitle, MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                OutputTableExportWorkflow.Run(
                    CreateElementConnectivityExportConfig(item, selectedTable, selectedObjectNames.Count),
                    _useCases,
                    _csiConnectionService,
                    _excelOutputService,
                    GetActiveOwnerWindow());
                StatusText = "Opened export options for " + selectedTable.Rows.Count.ToString(CultureInfo.InvariantCulture) + " selected " + item.Title + " row(s).";
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

        private static OutputTableExportConfig CreateElementConnectivityExportConfig(
            ElementConnectivityItem item,
            EtabsTableResult selectedTable,
            int selectedObjectCount)
        {
            string tableName = item == null || string.IsNullOrWhiteSpace(item.Title)
                ? "Object Connectivity"
                : item.Title;
            string groupName = item == null || string.IsNullOrWhiteSpace(item.GroupName)
                ? "Etabs Object Connectivity"
                : item.GroupName;
            int recordCount = selectedTable == null || selectedTable.Rows == null ? 0 : selectedTable.Rows.Count;

            return new OutputTableExportConfig
            {
                TableDisplayName = tableName,
                Breadcrumb = "ETABS Toolbox / Element Manipulation / " + groupName + " / " + tableName,
                Description = "Export " + tableName + " for " + selectedObjectCount.ToString(CultureInfo.InvariantCulture) + " selected ETABS object(s).",
                PopupProfileKey = "EtabsObjectConnectivity",
                EmptyDataMessage = "No " + tableName + " rows matched the current ETABS selection.",
                WorksheetNamePrefix = tableName,
                DefaultAddHeaders = true,
                StaticRecordCount = recordCount,
                StaticSuccessMessage = "Exported " + recordCount.ToString(CultureInfo.InvariantCulture) + " selected " + tableName + " row(s) to Excel.",
                StaticExportValuesFactory = addHeaders => CreateEtabsTableExportValues(selectedTable, addHeaders)
            };
        }

        private static HashSet<string> CreateSelectedObjectNameSet(IReadOnlyList<CsiSelectedObjectDto> selectedObjects)
        {
            var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            if (selectedObjects == null)
            {
                return names;
            }

            foreach (CsiSelectedObjectDto selectedObject in selectedObjects)
            {
                if (selectedObject == null || string.IsNullOrWhiteSpace(selectedObject.UniqueName))
                {
                    continue;
                }

                names.Add(selectedObject.UniqueName.Trim());
            }

            return names;
        }

        private static EtabsTableResult FilterConnectivityRowsBySelection(
            EtabsTableResult source,
            HashSet<string> selectedObjectNames)
        {
            var selectedTable = new EtabsTableResult
            {
                TableName = source == null ? string.Empty : source.TableName
            };

            if (source == null)
            {
                return selectedTable;
            }

            if (source.Headers != null)
            {
                selectedTable.Headers.AddRange(source.Headers);
            }

            if (source.Rows == null || selectedObjectNames == null || selectedObjectNames.Count == 0)
            {
                return selectedTable;
            }

            List<int> objectNameColumns = GetObjectNameColumnIndexes(source.Headers);
            AddMatchingConnectivityRows(source.Rows, selectedTable.Rows, selectedObjectNames, objectNameColumns);

            return selectedTable;
        }

        private static List<int> GetObjectNameColumnIndexes(IReadOnlyList<string> headers)
        {
            var indexes = new List<int>();
            if (headers == null)
            {
                return indexes;
            }

            for (int i = 0; i < headers.Count; i++)
            {
                string normalized = NormalizeConnectivityHeader(headers[i]);
                if (normalized == "UNIQUENAME" ||
                    normalized == "OBJECT" ||
                    normalized == "OBJECTNAME" ||
                    normalized == "OBJ" ||
                    normalized == "OBJNAME" ||
                    normalized == "NAME" ||
                    normalized == "LABEL")
                {
                    indexes.Add(i);
                }
            }

            return indexes;
        }

        private static string NormalizeConnectivityHeader(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return string.Empty;
            }

            var chars = new List<char>();
            foreach (char ch in value)
            {
                if (char.IsLetterOrDigit(ch))
                {
                    chars.Add(char.ToUpperInvariant(ch));
                }
            }

            return chars.Count == 0 ? string.Empty : new string(chars.ToArray());
        }

        private static void AddMatchingConnectivityRows(
            IList<List<string>> sourceRows,
            IList<List<string>> targetRows,
            HashSet<string> selectedObjectNames,
            IList<int> columnIndexes)
        {
            if (sourceRows == null || targetRows == null)
            {
                return;
            }

            foreach (List<string> row in sourceRows)
            {
                if (RowMatchesSelection(row, selectedObjectNames, columnIndexes))
                {
                    targetRows.Add(row == null ? new List<string>() : new List<string>(row));
                }
            }
        }

        private static bool RowMatchesSelection(
            IReadOnlyList<string> row,
            HashSet<string> selectedObjectNames,
            IList<int> columnIndexes)
        {
            if (row == null || selectedObjectNames == null || selectedObjectNames.Count == 0)
            {
                return false;
            }

            if (columnIndexes != null && columnIndexes.Count > 0)
            {
                foreach (int columnIndex in columnIndexes)
                {
                    if (columnIndex >= 0 && columnIndex < row.Count && CellMatchesSelection(row[columnIndex], selectedObjectNames))
                    {
                        return true;
                    }
                }

                return false;
            }

            for (int columnIndex = 0; columnIndex < row.Count; columnIndex++)
            {
                if (CellMatchesSelection(row[columnIndex], selectedObjectNames))
                {
                    return true;
                }
            }

            return false;
        }

        private static bool CellMatchesSelection(string cellValue, HashSet<string> selectedObjectNames)
        {
            return !string.IsNullOrWhiteSpace(cellValue) && selectedObjectNames.Contains(cellValue.Trim());
        }

        private static object[,] CreateEtabsTableExportValues(EtabsTableResult table, bool addHeaders)
        {
            int columnCount = GetEtabsTableExportColumnCount(table);
            int dataRowCount = table == null || table.Rows == null ? 0 : table.Rows.Count;
            int rowCount = dataRowCount + (addHeaders ? 1 : 0);
            if (columnCount == 0 || rowCount == 0)
            {
                return new object[0, 0];
            }

            var values = new object[rowCount, columnCount];
            int dataRowOffset = addHeaders ? 1 : 0;
            if (addHeaders && table != null && table.Headers != null)
            {
                for (int columnIndex = 0; columnIndex < columnCount; columnIndex++)
                {
                    values[0, columnIndex] = columnIndex < table.Headers.Count ? table.Headers[columnIndex] : string.Empty;
                }
            }

            if (table == null || table.Rows == null)
            {
                return values;
            }

            for (int rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++)
            {
                List<string> row = table.Rows[rowIndex];
                for (int columnIndex = 0; columnIndex < columnCount; columnIndex++)
                {
                    values[rowIndex + dataRowOffset, columnIndex] =
                        row != null && columnIndex < row.Count ? row[columnIndex] : string.Empty;
                }
            }

            return values;
        }

        private static int GetEtabsTableExportColumnCount(EtabsTableResult table)
        {
            if (table == null)
            {
                return 0;
            }

            int columnCount = table.Headers == null ? 0 : table.Headers.Count;
            if (table.Rows != null)
            {
                foreach (List<string> row in table.Rows)
                {
                    if (row != null && row.Count > columnCount)
                    {
                        columnCount = row.Count;
                    }
                }
            }

            return columnCount;
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
