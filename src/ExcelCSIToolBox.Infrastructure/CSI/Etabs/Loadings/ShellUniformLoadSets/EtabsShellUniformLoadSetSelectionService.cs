using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using ETABSv1;
using ExcelCSIToolBox.Application.Features.Loadings;
using ExcelCSIToolBox.Application.Interfaces.Etabs;
using ExcelCSIToolBox.Core.Abstractions.CSI;
using ExcelCSIToolBox.Core.Common.Results;
using ExcelCSIToolBox.Core.Contracts.CSI;
using ExcelCSIToolBox.Core.Tabular;
using ExcelCSIToolBox.Infrastructure.CSI.Common;

namespace ExcelCSIToolBox.Infrastructure.CSI.Etabs.Loadings.ShellUniformLoadSets
{
    public sealed class EtabsShellUniformLoadSetSelectionService : IEtabsShellUniformLoadSetSelectionService
    {
        private const string AssignmentTableKey = "Area Load Assignments - Uniform Load Sets";
        private readonly IEtabsConnectionService _connectionService;
        private readonly ICsiApiDispatcher _dispatcher;

        public EtabsShellUniformLoadSetSelectionService(
            IEtabsConnectionService connectionService,
            ICsiApiDispatcher dispatcher = null)
        {
            _connectionService = connectionService ?? throw new ArgumentNullException(nameof(connectionService));
            _dispatcher = dispatcher ?? new CurrentThreadCsiApiDispatcher();
        }

        public OperationResult<IReadOnlyList<string>> GetLoadSetNames()
        {
            return _dispatcher.Invoke(GetLoadSetNamesCore);
        }

        public OperationResult<IReadOnlyList<string>> GetStoryNames()
        {
            return _dispatcher.Invoke(GetStoryNamesCore);
        }

        public OperationResult<ShellUniformLoadSetSelectionResultDto> SelectShellsByLoadSets(IReadOnlyList<string> loadSetNames)
        {
            return SelectShellsByLoadSets(loadSetNames, (IReadOnlyList<string>)null);
        }

        public OperationResult<ShellUniformLoadSetSelectionResultDto> SelectShellsByLoadSets(
            IReadOnlyList<string> loadSetNames,
            string storyName)
        {
            IReadOnlyList<string> storyNames = string.IsNullOrWhiteSpace(storyName)
                ? null
                : new[] { storyName };
            return SelectShellsByLoadSets(loadSetNames, storyNames);
        }

        public OperationResult<ShellUniformLoadSetSelectionResultDto> SelectShellsByLoadSets(
            IReadOnlyList<string> loadSetNames,
            IReadOnlyList<string> storyNames)
        {
            return SelectShellsByLoadSets(loadSetNames, storyNames, null);
        }

        public OperationResult<ShellUniformLoadSetSelectionResultDto> SelectShellsByLoadSets(
            IReadOnlyList<string> loadSetNames,
            IReadOnlyList<string> storyNames,
            IProgress<ShellUniformLoadSetSelectionProgressDto> progress)
        {
            return _dispatcher.Invoke(() => SelectShellsByLoadSetsCore(loadSetNames, storyNames, progress));
        }

        private OperationResult<IReadOnlyList<string>> GetLoadSetNamesCore()
        {
            cSapModel sapModel;
            OperationResult modelResult = TryGetSapModel(out sapModel);
            if (!modelResult.IsSuccess)
            {
                return OperationResult<IReadOnlyList<string>>.Failure(modelResult.Message);
            }

            OperationResult<IReadOnlyList<string>> namesResult =
                EtabsShellUniformLoadSetTableService.GetDefinitionNames(sapModel);
            if (!namesResult.IsSuccess)
            {
                return OperationResult<IReadOnlyList<string>>.Failure(namesResult.Message);
            }

            IReadOnlyList<string> loadSetNames = ShellUniformLoadSetSelectionPlanner.NormalizeLoadSetNames(
                namesResult.Data ?? new string[0]);

            string message = loadSetNames.Count == 0
                ? "No Shell Uniform Load Sets exist in the connected ETABS model."
                : "Loaded " + loadSetNames.Count.ToString(CultureInfo.InvariantCulture) + " Shell Uniform Load Set(s).";
            return OperationResult<IReadOnlyList<string>>.Success(loadSetNames, message);
        }

        private OperationResult<IReadOnlyList<string>> GetStoryNamesCore()
        {
            cSapModel sapModel;
            OperationResult modelResult = TryGetSapModel(out sapModel);
            if (!modelResult.IsSuccess)
            {
                return OperationResult<IReadOnlyList<string>>.Failure(modelResult.Message);
            }

            try
            {
                int numberNames = 0;
                string[] names = null;
                int ret = sapModel.Story.GetNameList(ref numberNames, ref names);
                if (ret != 0)
                {
                    return OperationResult<IReadOnlyList<string>>.Failure(
                        "Failed to read ETABS stories (return code " + ret.ToString(CultureInfo.InvariantCulture) + ").");
                }

                IReadOnlyList<string> storyNames = NormalizeStoryNames(names);
                string message = storyNames.Count == 0
                    ? "No ETABS stories exist in the connected model."
                    : "Loaded " + storyNames.Count.ToString(CultureInfo.InvariantCulture) + " ETABS Story(s).";
                return OperationResult<IReadOnlyList<string>>.Success(storyNames, message);
            }
            catch (COMException ex)
            {
                return OperationResult<IReadOnlyList<string>>.Failure("ETABS COM error while reading stories: " + ex.Message);
            }
            catch (Exception ex)
            {
                return OperationResult<IReadOnlyList<string>>.Failure("Failed to read ETABS stories: " + ex.Message);
            }
        }

        private OperationResult<ShellUniformLoadSetSelectionResultDto> SelectShellsByLoadSetsCore(
            IReadOnlyList<string> rawLoadSetNames,
            IReadOnlyList<string> rawStoryNames,
            IProgress<ShellUniformLoadSetSelectionProgressDto> progress)
        {
            ReportProgress(progress, 0, 0, true, "Preparing selection...");
            IReadOnlyList<string> requestedLoadSetNames = ShellUniformLoadSetSelectionPlanner.NormalizeLoadSetNames(rawLoadSetNames);
            if (requestedLoadSetNames.Count == 0)
            {
                return OperationResult<ShellUniformLoadSetSelectionResultDto>.Failure("Select at least one Shell Uniform Load Set.");
            }

            IReadOnlyList<string> storyNames = NormalizeStoryNames(rawStoryNames);

            cSapModel sapModel;
            OperationResult modelResult = TryGetSapModel(out sapModel);
            if (!modelResult.IsSuccess)
            {
                return OperationResult<ShellUniformLoadSetSelectionResultDto>.Failure(modelResult.Message);
            }

            ReportProgress(progress, 0, 0, true, "Checking Shell Uniform Load Sets...");
            OperationResult<IReadOnlyList<string>> existingLoadSetsResult = GetLoadSetNamesCore();
            if (!existingLoadSetsResult.IsSuccess)
            {
                return OperationResult<ShellUniformLoadSetSelectionResultDto>.Failure(existingLoadSetsResult.Message);
            }

            IReadOnlyList<string> existingLoadSets = existingLoadSetsResult.Data ?? new string[0];
            if (existingLoadSets.Count == 0)
            {
                return OperationResult<ShellUniformLoadSetSelectionResultDto>.Failure("No Shell Uniform Load Sets exist in the connected ETABS model.");
            }

            ReportProgress(progress, 0, 0, true, "Reading Shell Uniform Load Set assignments...");
            OperationResult<IReadOnlyList<ShellUniformLoadSetAreaAssignmentDto>> assignmentsResult = ReadAreaLoadSetAssignments(sapModel);
            if (!assignmentsResult.IsSuccess)
            {
                return OperationResult<ShellUniformLoadSetSelectionResultDto>.Failure(assignmentsResult.Message);
            }

            ReportProgress(progress, 0, 0, true, "Resolving matching shell objects...");
            ShellUniformLoadSetSelectionPlan plan = ShellUniformLoadSetSelectionPlanner.CreatePlan(
                requestedLoadSetNames,
                existingLoadSets,
                assignmentsResult.Data,
                assignment => ResolveAreaObjectName(sapModel, assignment),
                storyNames);

            if (plan.AreaObjectNames.Count == 0)
            {
                return OperationResult<ShellUniformLoadSetSelectionResultDto>.Failure(CreateNoMatchingShellsMessage(plan));
            }

            OperationResult selectResult = SelectAreaObjects(sapModel, plan.AreaObjectNames, progress);
            if (!selectResult.IsSuccess)
            {
                return OperationResult<ShellUniformLoadSetSelectionResultDto>.Failure(selectResult.Message);
            }

            ShellUniformLoadSetSelectionResultDto result = CreateSelectionResult(plan);
            return OperationResult<ShellUniformLoadSetSelectionResultDto>.Success(result, result.Message);
        }

        private OperationResult TryGetSapModel(out cSapModel sapModel)
        {
            sapModel = _connectionService.SapModel as cSapModel;
            if (sapModel == null)
            {
                return OperationResult.Failure("The attached ETABS model is invalid. Please connect to ETABS and try again.");
            }

            return OperationResult.Success();
        }

        private static OperationResult<IReadOnlyList<ShellUniformLoadSetAreaAssignmentDto>> ReadAreaLoadSetAssignments(cSapModel sapModel)
        {
            try
            {
                int tableVersion = 0;
                string[] fieldKeyList = null;
                string[] fieldsKeysIncluded = null;
                int numberRecords = 0;
                string[] tableData = null;

                int ret = sapModel.DatabaseTables.GetTableForDisplayArray(
                    AssignmentTableKey,
                    ref fieldKeyList,
                    string.Empty,
                    ref tableVersion,
                    ref fieldsKeysIncluded,
                    ref numberRecords,
                    ref tableData);

                if (ret != 0)
                {
                    return OperationResult<IReadOnlyList<ShellUniformLoadSetAreaAssignmentDto>>.Failure(
                        "Could not read ETABS table \"" + AssignmentTableKey + "\" (return code " + ret.ToString(CultureInfo.InvariantCulture) + ").");
                }

                IReadOnlyList<string> fields = (fieldsKeysIncluded != null && fieldsKeysIncluded.Length > 0
                    ? fieldsKeysIncluded
                    : fieldKeyList) ?? new string[0];
                if (fields.Count == 0)
                {
                    return OperationResult<IReadOnlyList<ShellUniformLoadSetAreaAssignmentDto>>.Failure(
                        "The ETABS table \"" + AssignmentTableKey + "\" returned no field keys.");
                }

                int expectedLength = numberRecords * fields.Count;
                tableData = tableData ?? new string[0];
                if (tableData.Length < expectedLength)
                {
                    return OperationResult<IReadOnlyList<ShellUniformLoadSetAreaAssignmentDto>>.Failure(
                        "The ETABS table \"" + AssignmentTableKey + "\" returned inconsistent data length.");
                }

                int loadSetIndex = FindLoadSetFieldIndex(fields);
                int uniqueNameIndex = FindUniqueNameFieldIndex(fields);
                int storyIndex = FindStoryFieldIndex(fields);
                int labelIndex = FindAreaLabelFieldIndex(fields);
                OperationResult schemaResult = ValidateAssignmentSchema(loadSetIndex, uniqueNameIndex, storyIndex, labelIndex, fields);
                if (!schemaResult.IsSuccess)
                {
                    return OperationResult<IReadOnlyList<ShellUniformLoadSetAreaAssignmentDto>>.Failure(schemaResult.Message);
                }

                List<ShellUniformLoadSetAreaAssignmentDto> assignments = new List<ShellUniformLoadSetAreaAssignmentDto>();
                for (int recordIndex = 0; recordIndex < numberRecords; recordIndex++)
                {
                    assignments.Add(new ShellUniformLoadSetAreaAssignmentDto
                    {
                        Story = ReadTableValue(tableData, fields.Count, recordIndex, storyIndex),
                        Label = ReadTableValue(tableData, fields.Count, recordIndex, labelIndex),
                        UniqueName = ReadTableValue(tableData, fields.Count, recordIndex, uniqueNameIndex),
                        LoadSetName = ReadTableValue(tableData, fields.Count, recordIndex, loadSetIndex)
                    });
                }

                return OperationResult<IReadOnlyList<ShellUniformLoadSetAreaAssignmentDto>>.Success(assignments);
            }
            catch (COMException ex)
            {
                return OperationResult<IReadOnlyList<ShellUniformLoadSetAreaAssignmentDto>>.Failure(
                    "ETABS COM error while reading Shell Uniform Load Set assignments: " + ex.Message);
            }
            catch (Exception ex)
            {
                return OperationResult<IReadOnlyList<ShellUniformLoadSetAreaAssignmentDto>>.Failure(
                    "Failed to read Shell Uniform Load Set assignments: " + ex.Message);
            }
        }

        private static OperationResult ValidateAssignmentSchema(
            int loadSetIndex,
            int uniqueNameIndex,
            int storyIndex,
            int labelIndex,
            IReadOnlyList<string> fields)
        {
            if (loadSetIndex < 0)
            {
                return OperationResult.Failure(CreateMissingFieldMessage("load set", fields));
            }

            bool hasUniqueName = uniqueNameIndex >= 0;
            bool hasLabelAndStory = labelIndex >= 0 && storyIndex >= 0;
            if (!hasUniqueName && !hasLabelAndStory)
            {
                return OperationResult.Failure(
                    "The ETABS table \"" + AssignmentTableKey + "\" did not include either UniqueName or Story/Label fields required to resolve area objects.");
            }

            return OperationResult.Success();
        }

        private static string CreateMissingFieldMessage(string fieldDescription, IReadOnlyList<string> fields)
        {
            string fieldList = string.Join(", ", fields ?? new string[0]);
            string message = "The ETABS table \"" + AssignmentTableKey + "\" did not include a recognizable " + fieldDescription + " field. Detected fields: " + fieldList;
            return message;
        }

        private static int FindLoadSetFieldIndex(IReadOnlyList<string> fields)
        {
            return CsiTableFieldAliasResolver.FindFirstIndex(
                fields,
                "LoadSet",
                "Load Set",
                "Load Set Name",
                "LoadSetName",
                "UniformLoadSet",
                "Uniform Load Set",
                "UniformLoadSetName",
                "Uniform Load Set Name",
                "ShellUniformLoadSetName",
                "Shell Uniform Load Set Name");
        }

        private static int FindUniqueNameFieldIndex(IReadOnlyList<string> fields)
        {
            return CsiTableFieldAliasResolver.FindFirstIndex(fields, "UniqueName", "Unique Name", "ObjectName", "Object Name");
        }

        private static int FindStoryFieldIndex(IReadOnlyList<string> fields)
        {
            return CsiTableFieldAliasResolver.FindFirstIndex(fields, "Story", "StoryName", "Story Name");
        }

        private static int FindAreaLabelFieldIndex(IReadOnlyList<string> fields)
        {
            return CsiTableFieldAliasResolver.FindFirstIndex(
                fields,
                "Label",
                "LabelName",
                "Label Name",
                "Area",
                "AreaName",
                "Area Name",
                "Shell",
                "ShellName",
                "Shell Name");
        }

        private static string ReadTableValue(string[] tableData, int fieldCount, int recordIndex, int fieldIndex)
        {
            if (fieldIndex < 0 || tableData == null || fieldCount <= 0)
            {
                return string.Empty;
            }

            int dataIndex = recordIndex * fieldCount + fieldIndex;
            string value = dataIndex >= 0 && dataIndex < tableData.Length ? tableData[dataIndex] : string.Empty;
            return value ?? string.Empty;
        }

        private static string ResolveAreaObjectName(cSapModel sapModel, ShellUniformLoadSetAreaAssignmentDto assignment)
        {
            if (assignment == null)
            {
                return string.Empty;
            }

            string uniqueName = NormalizeName(assignment.UniqueName);
            if (!string.IsNullOrWhiteSpace(uniqueName))
            {
                return uniqueName;
            }

            string label = NormalizeName(assignment.Label);
            string story = NormalizeName(assignment.Story);
            if (string.IsNullOrWhiteSpace(label) || string.IsNullOrWhiteSpace(story))
            {
                return string.Empty;
            }

            string areaObjectName = string.Empty;
            int ret = sapModel.AreaObj.GetNameFromLabel(label, story, ref areaObjectName);
            return ret == 0 ? NormalizeName(areaObjectName) : string.Empty;
        }

        private static OperationResult SelectAreaObjects(
            cSapModel sapModel,
            IReadOnlyList<string> areaObjectNames,
            IProgress<ShellUniformLoadSetSelectionProgressDto> progress)
        {
            int total = areaObjectNames == null ? 0 : areaObjectNames.Count;
            ReportProgress(progress, 0, total, false, "Clearing current ETABS selection...");
            int clearRet = sapModel.SelectObj.ClearSelection();
            if (clearRet != 0)
            {
                return OperationResult.Failure("Failed to clear the current ETABS selection (return code " + clearRet.ToString(CultureInfo.InvariantCulture) + ").");
            }

            List<string> failures = new List<string>();
            int current = 0;
            foreach (string areaObjectName in areaObjectNames ?? new string[0])
            {
                int ret = sapModel.AreaObj.SetSelected(areaObjectName, true, eItemType.Objects);
                current++;
                ReportProgress(
                    progress,
                    current,
                    total,
                    false,
                    "Selecting ETABS shell objects: " +
                    current.ToString(CultureInfo.InvariantCulture) +
                    " / " +
                    total.ToString(CultureInfo.InvariantCulture));
                if (ret != 0)
                {
                    failures.Add(areaObjectName + " (return code " + ret.ToString(CultureInfo.InvariantCulture) + ")");
                }
            }

            if (failures.Count > 0)
            {
                return OperationResult.Failure(
                    "ETABS could not select " + failures.Count.ToString(CultureInfo.InvariantCulture) + " area object(s): " +
                    string.Join("; ", failures.Take(10)) +
                    (failures.Count > 10 ? "; ..." : string.Empty));
            }

            ReportProgress(progress, total, total, false, "Refreshing ETABS view...");
            int refreshRet = sapModel.View.RefreshView(0, false);
            if (refreshRet != 0)
            {
                return OperationResult.Failure("ETABS selected the shells, but View.RefreshView failed (return code " + refreshRet.ToString(CultureInfo.InvariantCulture) + ").");
            }

            ReportProgress(progress, total, total, false, "Selection complete.");
            return OperationResult.Success();
        }

        private static ShellUniformLoadSetSelectionResultDto CreateSelectionResult(ShellUniformLoadSetSelectionPlan plan)
        {
            List<string> warnings = new List<string>();
            if (plan.UnknownLoadSetNames.Count > 0)
            {
                warnings.Add(plan.UnknownLoadSetNames.Count.ToString(CultureInfo.InvariantCulture) + " selected load set(s) do not exist in the connected ETABS model.");
            }

            if (plan.UnresolvedAreaCount > 0)
            {
                warnings.Add(plan.UnresolvedAreaCount.ToString(CultureInfo.InvariantCulture) + " area assignment(s) could not be resolved to ETABS objects.");
            }

            string storyText = CreateSelectedStoryMessageFragment(plan.SelectedStoryNames);
            string message = "Selected " + plan.AreaObjectNames.Count.ToString(CultureInfo.InvariantCulture) +
                " shell(s) assigned to " + plan.MatchedLoadSetNames.Count.ToString(CultureInfo.InvariantCulture) + " load set(s)" +
                storyText + ".";

            ShellUniformLoadSetSelectionResultDto result = new ShellUniformLoadSetSelectionResultDto
            {
                RequestedLoadSetCount = plan.RequestedLoadSetNames.Count,
                MatchedLoadSetCount = plan.MatchedLoadSetNames.Count,
                SelectedShellCount = plan.AreaObjectNames.Count,
                UnresolvedAreaCount = plan.UnresolvedAreaCount,
                DuplicateShellCount = plan.DuplicateShellCount,
                UnknownLoadSetCount = plan.UnknownLoadSetNames.Count,
                SelectedStoryName = plan.SelectedStoryName,
                SelectedStoryNames = plan.SelectedStoryNames.ToList(),
                Message = message,
                WarningMessage = string.Join(" ", warnings),
                SelectedShellNames = plan.AreaObjectNames.ToList(),
                UnknownLoadSetNames = plan.UnknownLoadSetNames.ToList(),
                UnresolvedAreaReferences = plan.UnresolvedAreaReferences.ToList()
            };

            return result;
        }

        private static string CreateNoMatchingShellsMessage(ShellUniformLoadSetSelectionPlan plan)
        {
            if (plan.UnknownLoadSetNames.Count == plan.RequestedLoadSetNames.Count)
            {
                return "None of the selected Shell Uniform Load Sets exist in the connected ETABS model.";
            }

            if (plan.MatchingAssignmentCount == 0)
            {
                if (plan.SelectedStoryNames.Count > 0)
                {
                    return "No shells are assigned to the selected load sets" +
                        CreateSelectedStoryMessageFragment(plan.SelectedStoryNames) + ".";
                }

                return "No shells are assigned to the selected load sets.";
            }

            if (plan.UnresolvedAreaCount > 0)
            {
                return "No matching shell could be resolved to ETABS area objects. " +
                    plan.UnresolvedAreaCount.ToString(CultureInfo.InvariantCulture) +
                    " area assignment(s) could not be resolved.";
            }

            return "No shells are assigned to the selected load sets.";
        }

        private static string NormalizeName(string value)
        {
            string name = string.IsNullOrWhiteSpace(value) ? string.Empty : value.Trim();
            return name;
        }

        private static string CreateSelectedStoryMessageFragment(IReadOnlyList<string> selectedStoryNames)
        {
            if (selectedStoryNames == null || selectedStoryNames.Count == 0)
            {
                return string.Empty;
            }

            if (selectedStoryNames.Count == 1)
            {
                return " on story '" + selectedStoryNames[0] + "'";
            }

            return " on " + selectedStoryNames.Count.ToString(CultureInfo.InvariantCulture) + " selected story(s)";
        }

        private static void ReportProgress(
            IProgress<ShellUniformLoadSetSelectionProgressDto> progress,
            int current,
            int total,
            bool isIndeterminate,
            string message)
        {
            if (progress == null)
            {
                return;
            }

            progress.Report(new ShellUniformLoadSetSelectionProgressDto
            {
                Current = current,
                Total = total,
                IsIndeterminate = isIndeterminate,
                Message = message
            });
        }

        private static IReadOnlyList<string> NormalizeStoryNames(IEnumerable<string> storyNames)
        {
            HashSet<string> seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            List<string> names = new List<string>();
            foreach (string rawName in storyNames ?? new string[0])
            {
                string name = NormalizeName(rawName);
                if (string.IsNullOrWhiteSpace(name) || !seen.Add(name))
                {
                    continue;
                }

                names.Add(name);
            }

            return names;
        }
    }
}
