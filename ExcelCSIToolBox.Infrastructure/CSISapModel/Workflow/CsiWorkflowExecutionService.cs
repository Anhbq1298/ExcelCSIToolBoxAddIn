using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text.RegularExpressions;
using ExcelCSIToolBox.Core.Abstractions.CSI;
using ExcelCSIToolBox.Core.Common.Results;
using ExcelCSIToolBox.Data.CSISapModel.FrameObject;
using ExcelCSIToolBox.Data.CSISapModel.Workflow;
using ExcelCSIToolBox.Data.DTOs.CSI;
using ExcelCSIToolBox.Data.Models;

namespace ExcelCSIToolBox.Infrastructure.CSISapModel.Workflow
{
    public sealed class CsiWorkflowExecutionService
    {
        public OperationResult<CsiWorkflowResultDto> Execute(
            ICSISapModelConnectionService service,
            CsiWorkflowRequestDto request)
        {
            if (service == null)
            {
                return OperationResult<CsiWorkflowResultDto>.Failure("active CSI model is not available.");
            }

            string userInput = request == null ? null : request.UserInput;
            List<CsiTaskDto> tasks = request != null && request.PlannedTasks != null && request.PlannedTasks.Count > 0
                ? NormalizePlannedTasks(request.PlannedTasks)
                : ParseTasks(userInput);
            var result = new CsiWorkflowResultDto
            {
                TotalTasksDetected = tasks.Count,
                Results = new List<CsiTaskResultDto>()
            };

            var succeededTaskIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var failedOrSkippedTaskIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var knownFrames = new List<string>();

            foreach (CsiTaskDto task in tasks)
            {
                string dependencyFailure = GetDependencyFailure(task, failedOrSkippedTaskIds);
                if (!string.IsNullOrWhiteSpace(dependencyFailure))
                {
                    AddResult(result, new CsiTaskResultDto
                    {
                        TaskId = task.TaskId,
                        TaskType = task.TaskType,
                        Operation = task.Operation,
                        Skipped = true,
                        FailureReason = dependencyFailure
                    });
                    failedOrSkippedTaskIds.Add(task.TaskId);
                    continue;
                }

                CsiTaskResultDto taskResult = ExecuteTask(service, task, knownFrames);
                AddResult(result, taskResult);
                if (taskResult.Success)
                {
                    succeededTaskIds.Add(task.TaskId);
                    if (IsFrameCreation(task) && !string.IsNullOrWhiteSpace(taskResult.ObjectName))
                    {
                        AddUnique(knownFrames, taskResult.ObjectName);
                    }
                }
                else
                {
                    failedOrSkippedTaskIds.Add(task.TaskId);
                }
            }

            return OperationResult<CsiWorkflowResultDto>.Success(result);
        }

        private static List<CsiTaskDto> NormalizePlannedTasks(List<CsiTaskDto> plannedTasks)
        {
            var tasks = new List<CsiTaskDto>();
            if (plannedTasks == null)
            {
                return tasks;
            }

            foreach (CsiTaskDto plannedTask in plannedTasks)
            {
                if (plannedTask == null)
                {
                    continue;
                }

                tasks.Add(new CsiTaskDto
                {
                    TaskId = plannedTask.TaskId,
                    TaskType = plannedTask.TaskType,
                    Operation = plannedTask.Operation,
                    Arguments = plannedTask.Arguments ?? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase),
                    DependsOn = plannedTask.DependsOn ?? new List<string>()
                });
            }

            ApplyTaskIds(tasks);
            ApplyDependencies(tasks);
            return tasks;
        }

        private static List<CsiTaskDto> ParseTasks(string userInput)
        {
            var tasks = new List<CsiTaskDto>();
            if (string.IsNullOrWhiteSpace(userInput))
            {
                tasks.Add(CreateTask("Query", "InvalidRequest", null));
                tasks[0].Arguments["failureReason"] = "User input is required.";
                return tasks;
            }

            string text = userInput.Trim();
            ParsePointTasks(text, tasks);
            ParseFrameTasks(text, tasks);
            ParseFrameSectionTask(text, tasks);
            ParseFrameDistributedLoadTask(text, tasks);
            ParseFrameLengthQueryTask(text, tasks);
            ParseSelectionTask(text, tasks);
            ParseUnsupportedRecognizedTasks(text, tasks);

            if (tasks.Count == 0)
            {
                tasks.Add(CreateTask("Query", "Unsupported", null));
                tasks[0].Arguments["failureReason"] = "No supported CSI workflow task was detected.";
            }

            ApplyDependencies(tasks);
            return tasks;
        }

        private static void ParsePointTasks(string text, List<CsiTaskDto> tasks)
        {
            MatchCollection matches = Regex.Matches(
                text,
                @"\b(?<name>[A-Za-z_][A-Za-z0-9_\-\.]*)\s+at\s+(?<x>-?\d+(?:\.\d+)?)\s*,\s*(?<y>-?\d+(?:\.\d+)?)\s*,\s*(?<z>-?\d+(?:\.\d+)?)",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

            foreach (Match match in matches)
            {
                string name = match.Groups["name"].Value;
                if (IsReservedName(name))
                {
                    continue;
                }

                CsiTaskDto task = CreateTask("PointObj", "AddCartesian", null);
                task.Arguments["name"] = name;
                task.Arguments["x"] = match.Groups["x"].Value;
                task.Arguments["y"] = match.Groups["y"].Value;
                task.Arguments["z"] = match.Groups["z"].Value;
                tasks.Add(task);
            }
        }

        private static void ParseFrameTasks(string text, List<CsiTaskDto> tasks)
        {
            MatchCollection byPointMatches = Regex.Matches(
                text,
                @"\b(?:add|create|connect|draw)?\s*(?:beam|frame)?\s*(?<frame>[A-Za-z_][A-Za-z0-9_\-\.]*)?\s*(?:between|from)\s+(?<pi>[A-Za-z_][A-Za-z0-9_\-\.]*)\s+(?:and|to)\s+(?<pj>[A-Za-z_][A-Za-z0-9_\-\.]*)",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

            foreach (Match match in byPointMatches)
            {
                string frameName = CleanObjectName(match.Groups["frame"].Value);
                if (IsReservedName(frameName))
                {
                    frameName = null;
                }

                CsiTaskDto task = CreateTask("FrameObj", "Add", null);
                task.Arguments["userName"] = frameName;
                task.Arguments["pointIName"] = match.Groups["pi"].Value;
                task.Arguments["pointJName"] = match.Groups["pj"].Value;
                task.Arguments["propName"] = ExtractSectionName(text);
                tasks.Add(task);
            }

            if (byPointMatches.Count > 0)
            {
                return;
            }

            MatchCollection coordinateMatches = Regex.Matches(
                text,
                @"(-?\d+(?:\.\d+)?)\s*,\s*(-?\d+(?:\.\d+)?)\s*,\s*(-?\d+(?:\.\d+)?)",
                RegexOptions.CultureInvariant);
            if (Regex.IsMatch(text, @"\b(frame|beam)\b", RegexOptions.IgnoreCase) && coordinateMatches.Count >= 2)
            {
                CsiTaskDto task = CreateTask("FrameObj", "Add", null);
                string frameName = ExtractNamedObject(text, @"\b(?:frame|beam)\s+(?<name>[A-Za-z_][A-Za-z0-9_\-\.]*)");
                task.Arguments["userName"] = frameName;
                task.Arguments["xi"] = coordinateMatches[0].Groups[1].Value;
                task.Arguments["yi"] = coordinateMatches[0].Groups[2].Value;
                task.Arguments["zi"] = coordinateMatches[0].Groups[3].Value;
                task.Arguments["xj"] = coordinateMatches[1].Groups[1].Value;
                task.Arguments["yj"] = coordinateMatches[1].Groups[2].Value;
                task.Arguments["zj"] = coordinateMatches[1].Groups[3].Value;
                task.Arguments["propName"] = ExtractSectionName(text);
                tasks.Add(task);
            }
        }

        private static void ParseFrameSectionTask(string text, List<CsiTaskDto> tasks)
        {
            string sectionName = ExtractSectionName(text);
            if (string.IsNullOrWhiteSpace(sectionName) ||
                !Regex.IsMatch(text, @"\b(assign|set|update|apply)\b.*\b(section|property)\b|\b(section|property)\b.*\b(assign|set|update|apply)\b", RegexOptions.IgnoreCase))
            {
                return;
            }

            CsiTaskDto task = CreateTask("FrameObj", "AssignSection", null);
            task.Arguments["sectionName"] = sectionName;
            task.Arguments["frameNames"] = string.Join(",", ExtractFrameNames(text));
            tasks.Add(task);
        }

        private static void ParseFrameDistributedLoadTask(string text, List<CsiTaskDto> tasks)
        {
            if (!Regex.IsMatch(text, @"\b(udl|distributed|uniform|kN/m|kn/m)\b", RegexOptions.IgnoreCase))
            {
                return;
            }

            Match valueMatch = Regex.Match(text, @"(?<value>-?\d+(?:\.\d+)?)\s*(?:kN/m|kn/m)", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
            CsiTaskDto task = CreateTask("FrameLoad", "AssignDistributed", null);
            task.Arguments["frameNames"] = string.Join(",", ExtractFrameNames(text));
            task.Arguments["loadPattern"] = ExtractLoadPattern(text);
            task.Arguments["direction"] = Regex.IsMatch(text, @"\b(down|downward|gravity)\b", RegexOptions.IgnoreCase) ? "6" : "6";
            task.Arguments["value1"] = valueMatch.Success ? "-" + valueMatch.Groups["value"].Value.TrimStart('-') : null;
            task.Arguments["value2"] = valueMatch.Success ? "-" + valueMatch.Groups["value"].Value.TrimStart('-') : null;
            tasks.Add(task);
        }

        private static void ParseFrameLengthQueryTask(string text, List<CsiTaskDto> tasks)
        {
            if (!Regex.IsMatch(text, @"\b(get|extract|query|read|show)\b.*\b(frame|beam).*\b(length|lengths)\b|\b(length|lengths)\b.*\b(frame|beam)\b", RegexOptions.IgnoreCase))
            {
                return;
            }

            CsiTaskDto task = CreateTask("Query", "ExtractFrameLengths", null);
            task.Arguments["frameNames"] = string.Join(",", ExtractFrameNames(text));
            tasks.Add(task);
        }

        private static void ParseSelectionTask(string text, List<CsiTaskDto> tasks)
        {
            if (!Regex.IsMatch(text, @"\bselect\b", RegexOptions.IgnoreCase))
            {
                return;
            }

            List<string> frameNames = ExtractFrameNames(text);
            if (frameNames.Count == 0)
            {
                return;
            }

            CsiTaskDto task = CreateTask("Selection", "SelectFrames", null);
            task.Arguments["frameNames"] = string.Join(",", frameNames);
            tasks.Add(task);
        }

        private static void ParseUnsupportedRecognizedTasks(string text, List<CsiTaskDto> tasks)
        {
            AddUnsupportedIfDetected(text, tasks, "AreaObj", @"\b(area|shell|slab|wall)\b");
            AddUnsupportedIfDetected(text, tasks, "LoadPattern", @"\bload pattern\b");
            AddUnsupportedIfDetected(text, tasks, "LoadCase", @"\bload case\b");
            AddUnsupportedIfDetected(text, tasks, "LoadCombination", @"\b(load combination|combo)\b");
        }

        private static void AddUnsupportedIfDetected(string text, List<CsiTaskDto> tasks, string taskType, string pattern)
        {
            if (!Regex.IsMatch(text, pattern, RegexOptions.IgnoreCase))
            {
                return;
            }

            foreach (CsiTaskDto existing in tasks)
            {
                if (string.Equals(existing.TaskType, taskType, StringComparison.OrdinalIgnoreCase))
                {
                    return;
                }
            }

            CsiTaskDto task = CreateTask(taskType, "Unsupported", null);
            task.Arguments["failureReason"] = taskType + " workflow parsing is not implemented for this request.";
            tasks.Add(task);
        }

        private static void ApplyDependencies(List<CsiTaskDto> tasks)
        {
            string lastPointTask = LastTaskId(tasks, "PointObj", "AddCartesian");
            string lastFrameTask = LastTaskId(tasks, "FrameObj", "Add");

            foreach (CsiTaskDto task in tasks)
            {
                if (string.Equals(task.TaskType, "FrameObj", StringComparison.OrdinalIgnoreCase) &&
                    string.Equals(task.Operation, "Add", StringComparison.OrdinalIgnoreCase) &&
                    !string.IsNullOrWhiteSpace(lastPointTask) &&
                    TaskIndex(tasks, lastPointTask) < TaskIndex(tasks, task.TaskId))
                {
                    AddDependency(task, lastPointTask);
                }

                if ((string.Equals(task.Operation, "AssignSection", StringComparison.OrdinalIgnoreCase) ||
                     string.Equals(task.Operation, "AssignDistributed", StringComparison.OrdinalIgnoreCase) ||
                     string.Equals(task.Operation, "ExtractFrameLengths", StringComparison.OrdinalIgnoreCase)) &&
                    !string.IsNullOrWhiteSpace(lastFrameTask) &&
                    TaskIndex(tasks, lastFrameTask) < TaskIndex(tasks, task.TaskId))
                {
                    AddDependency(task, lastFrameTask);
                }
            }
        }

        private static CsiTaskResultDto ExecuteTask(
            ICSISapModelConnectionService service,
            CsiTaskDto task,
            List<string> knownFrames)
        {
            string failureReason;
            if (TryGetArgument(task, "failureReason", out failureReason))
            {
                return Failed(task, failureReason);
            }

            if (task.TaskType == "PointObj" && task.Operation == "AddCartesian")
            {
                return ExecuteAddPoint(service, task);
            }

            if (task.TaskType == "FrameObj" && task.Operation == "Add")
            {
                return ExecuteAddFrame(service, task);
            }

            if (task.TaskType == "FrameObj" && task.Operation == "AssignSection")
            {
                return ExecuteAssignSection(service, task, knownFrames);
            }

            if (task.TaskType == "FrameLoad" && task.Operation == "AssignDistributed")
            {
                return ExecuteAssignDistributedLoad(service, task, knownFrames);
            }

            if (task.TaskType == "Query" && task.Operation == "ExtractFrameLengths")
            {
                return ExecuteExtractFrameLengths(service, task, knownFrames);
            }

            if (task.TaskType == "Selection" && task.Operation == "SelectFrames")
            {
                return ExecuteSelectFrames(service, task, knownFrames);
            }

            return Failed(task, task.TaskType + " " + task.Operation + " is not supported by the workflow executor.");
        }

        private static CsiTaskResultDto ExecuteAddPoint(ICSISapModelConnectionService service, CsiTaskDto task)
        {
            string name = Get(task, "name");
            double x;
            double y;
            double z;
            if (!TryParseDouble(Get(task, "x"), out x) ||
                !TryParseDouble(Get(task, "y"), out y) ||
                !TryParseDouble(Get(task, "z"), out z))
            {
                return Failed(task, "Point coordinates are invalid.");
            }

            OperationResult<CSISapModelAddPointsResultDTO> result = service.AddPointsByCartesian(new[]
            {
                new CSISapModelPointCartesianInput
                {
                    ExcelRowNumber = 1,
                    UniqueName = name,
                    X = x,
                    Y = y,
                    Z = z
                }
            });

            if (!result.IsSuccess || result.Data == null || result.Data.AddedCount <= 0)
            {
                return Failed(task, result.Message ?? "PointObj.AddCartesian returned without adding a point.");
            }

            return Succeeded(task, name, "Success");
        }

        private static CsiTaskResultDto ExecuteAddFrame(ICSISapModelConnectionService service, CsiTaskDto task)
        {
            var request = new FrameAddRequestDto
            {
                UserName = Get(task, "userName"),
                PropName = Get(task, "propName"),
                PointIName = Get(task, "pointIName"),
                PointJName = Get(task, "pointJName"),
                Xi = ParseNullableDouble(Get(task, "xi")),
                Yi = ParseNullableDouble(Get(task, "yi")),
                Zi = ParseNullableDouble(Get(task, "zi")),
                Xj = ParseNullableDouble(Get(task, "xj")),
                Yj = ParseNullableDouble(Get(task, "yj")),
                Zj = ParseNullableDouble(Get(task, "zj"))
            };

            OperationResult<FrameAddBatchResultDto> result = service.AddFrameObjects(new FrameAddBatchRequestDto
            {
                Frames = new List<FrameAddRequestDto> { request }
            });

            if (!result.IsSuccess || result.Data == null || result.Data.Results.Count == 0)
            {
                return Failed(task, result.Message ?? "Frame add failed.");
            }

            FrameAddResultDto frameResult = result.Data.Results[0];
            if (!frameResult.Success)
            {
                return Failed(task, frameResult.FailureReason ?? "Frame add failed.");
            }

            return Succeeded(task, frameResult.FrameName, "Success");
        }

        private static CsiTaskResultDto ExecuteAssignSection(
            ICSISapModelConnectionService service,
            CsiTaskDto task,
            List<string> knownFrames)
        {
            string sectionName = Get(task, "sectionName");
            List<string> frameNames = GetTargetFrames(task, knownFrames);
            if (frameNames.Count == 0)
            {
                return Failed(task, "At least one frame name is required.");
            }

            OperationResult result = service.AssignFrameSection(frameNames, sectionName);
            return result.IsSuccess
                ? Succeeded(task, string.Join(",", frameNames), "Success")
                : Failed(task, result.Message);
        }

        private static CsiTaskResultDto ExecuteAssignDistributedLoad(
            ICSISapModelConnectionService service,
            CsiTaskDto task,
            List<string> knownFrames)
        {
            string loadPattern = Get(task, "loadPattern");
            if (string.IsNullOrWhiteSpace(loadPattern))
            {
                return Failed(task, "Load pattern is required.");
            }

            double value1;
            double value2;
            int direction;
            if (!TryParseDouble(Get(task, "value1"), out value1) ||
                !TryParseDouble(Get(task, "value2"), out value2) ||
                !int.TryParse(Get(task, "direction"), out direction))
            {
                return Failed(task, "Distributed load value or direction is invalid.");
            }

            List<string> frameNames = GetTargetFrames(task, knownFrames);
            if (frameNames.Count == 0)
            {
                return Failed(task, "At least one frame name is required.");
            }

            OperationResult result = service.AssignFrameDistributedLoad(frameNames, loadPattern, direction, value1, value2);
            return result.IsSuccess
                ? Succeeded(task, string.Join(",", frameNames), "Success")
                : Failed(task, result.Message);
        }

        private static CsiTaskResultDto ExecuteExtractFrameLengths(
            ICSISapModelConnectionService service,
            CsiTaskDto task,
            List<string> knownFrames)
        {
            List<string> frameNames = GetTargetFrames(task, knownFrames);
            if (frameNames.Count == 0)
            {
                OperationResult<IReadOnlyList<string>> allFrames = service.GetFrameNames();
                if (allFrames.IsSuccess && allFrames.Data != null)
                {
                    frameNames.AddRange(allFrames.Data);
                }
            }

            if (frameNames.Count == 0)
            {
                return Failed(task, "At least one frame name is required.");
            }

            var lengths = new List<string>();
            foreach (string frameName in frameNames)
            {
                OperationResult<FrameEndPointInfo> pointsResult = service.GetFramePoints(frameName);
                if (!pointsResult.IsSuccess)
                {
                    return Failed(task, pointsResult.Message);
                }

                OperationResult<ExcelCSIToolBox.Data.CSISapModel.PointObject.PointObjectInfo> pi = service.GetPointCoordinates(pointsResult.Data.PointI);
                OperationResult<ExcelCSIToolBox.Data.CSISapModel.PointObject.PointObjectInfo> pj = service.GetPointCoordinates(pointsResult.Data.PointJ);
                if (!pi.IsSuccess || !pj.IsSuccess)
                {
                    return Failed(task, "Failed to read frame endpoint coordinates.");
                }

                double dx = pj.Data.X - pi.Data.X;
                double dy = pj.Data.Y - pi.Data.Y;
                double dz = pj.Data.Z - pi.Data.Z;
                double length = Math.Sqrt(dx * dx + dy * dy + dz * dz);
                lengths.Add(frameName + "=" + length.ToString("0.###", CultureInfo.InvariantCulture));
            }

            return Succeeded(task, string.Join("; ", lengths), "Success");
        }

        private static CsiTaskResultDto ExecuteSelectFrames(
            ICSISapModelConnectionService service,
            CsiTaskDto task,
            List<string> knownFrames)
        {
            List<string> frameNames = GetTargetFrames(task, knownFrames);
            if (frameNames.Count == 0)
            {
                return Failed(task, "At least one frame name is required.");
            }

            OperationResult result = service.SelectFramesByUniqueNames(frameNames);
            return result.IsSuccess
                ? Succeeded(task, string.Join(",", frameNames), "Success")
                : Failed(task, result.Message);
        }

        private static List<string> GetTargetFrames(CsiTaskDto task, List<string> knownFrames)
        {
            var result = SplitNames(Get(task, "frameNames"));
            if (result.Count == 0 && knownFrames != null)
            {
                result.AddRange(knownFrames);
            }

            return result;
        }

        private static string ExtractSectionName(string text)
        {
            string section = ExtractNamedObject(text, @"\bsection\s+(?<name>[A-Za-z_][A-Za-z0-9_\-\.]*)");
            if (!string.IsNullOrWhiteSpace(section))
            {
                return section;
            }

            return ExtractNamedObject(text, @"\bassign\s+(?<name>[A-Za-z_][A-Za-z0-9_\-\.]*)\s+section\b");
        }

        private static string ExtractLoadPattern(string text)
        {
            string pattern = ExtractNamedObject(text, @"\b(?:load\s+pattern|pattern|case)\s+(?<name>[A-Za-z_][A-Za-z0-9_\-\.]*)");
            if (!string.IsNullOrWhiteSpace(pattern))
            {
                return pattern;
            }

            return null;
        }

        private static List<string> ExtractFrameNames(string text)
        {
            var names = new List<string>();
            MatchCollection matches = Regex.Matches(text, @"\b(?:frame|beam)\s+(?<name>[A-Za-z_][A-Za-z0-9_\-\.]*)", RegexOptions.IgnoreCase);
            foreach (Match match in matches)
            {
                string name = CleanObjectName(match.Groups["name"].Value);
                if (!IsReservedName(name))
                {
                    AddUnique(names, name);
                }
            }

            return names;
        }

        private static string ExtractNamedObject(string text, string pattern)
        {
            Match match = Regex.Match(text ?? string.Empty, pattern, RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
            return match.Success ? CleanObjectName(match.Groups["name"].Value) : null;
        }

        private static CsiTaskDto CreateTask(string taskType, string operation, List<string> dependsOn)
        {
            return new CsiTaskDto
            {
                TaskId = null,
                TaskType = taskType,
                Operation = operation,
                Arguments = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase),
                DependsOn = dependsOn ?? new List<string>()
            };
        }

        private static void ApplyTaskIds(List<CsiTaskDto> tasks)
        {
            for (int i = 0; i < tasks.Count; i++)
            {
                if (string.IsNullOrWhiteSpace(tasks[i].TaskId))
                {
                    tasks[i].TaskId = "task-" + (i + 1).ToString(CultureInfo.InvariantCulture);
                }
            }
        }

        private static string LastTaskId(List<CsiTaskDto> tasks, string taskType, string operation)
        {
            ApplyTaskIds(tasks);
            for (int i = tasks.Count - 1; i >= 0; i--)
            {
                if (string.Equals(tasks[i].TaskType, taskType, StringComparison.OrdinalIgnoreCase) &&
                    string.Equals(tasks[i].Operation, operation, StringComparison.OrdinalIgnoreCase))
                {
                    return tasks[i].TaskId;
                }
            }

            return null;
        }

        private static int TaskIndex(List<CsiTaskDto> tasks, string taskId)
        {
            for (int i = 0; i < tasks.Count; i++)
            {
                if (string.Equals(tasks[i].TaskId, taskId, StringComparison.OrdinalIgnoreCase))
                {
                    return i;
                }
            }

            return -1;
        }

        private static void AddDependency(CsiTaskDto task, string dependencyTaskId)
        {
            if (task.DependsOn == null)
            {
                task.DependsOn = new List<string>();
            }

            if (!task.DependsOn.Contains(dependencyTaskId))
            {
                task.DependsOn.Add(dependencyTaskId);
            }
        }

        private static string GetDependencyFailure(CsiTaskDto task, HashSet<string> failedOrSkippedTaskIds)
        {
            if (task.DependsOn == null)
            {
                return null;
            }

            foreach (string dependency in task.DependsOn)
            {
                if (failedOrSkippedTaskIds.Contains(dependency))
                {
                    return "Dependency failed or was skipped: " + dependency;
                }
            }

            return null;
        }

        private static void AddResult(CsiWorkflowResultDto workflowResult, CsiTaskResultDto taskResult)
        {
            workflowResult.Results.Add(taskResult);
            if (taskResult.Skipped)
            {
                workflowResult.Skipped++;
            }
            else if (taskResult.Success)
            {
                workflowResult.Succeeded++;
            }
            else
            {
                workflowResult.Failed++;
            }
        }

        private static CsiTaskResultDto Succeeded(CsiTaskDto task, string objectName, string message)
        {
            return new CsiTaskResultDto
            {
                TaskId = task.TaskId,
                TaskType = task.TaskType,
                Operation = task.Operation,
                Success = true,
                ObjectName = objectName,
                Message = message
            };
        }

        private static CsiTaskResultDto Failed(CsiTaskDto task, string reason)
        {
            return new CsiTaskResultDto
            {
                TaskId = task.TaskId,
                TaskType = task.TaskType,
                Operation = task.Operation,
                Success = false,
                FailureReason = reason
            };
        }

        private static bool TryGetArgument(CsiTaskDto task, string key, out string value)
        {
            value = null;
            return task.Arguments != null && task.Arguments.TryGetValue(key, out value) && !string.IsNullOrWhiteSpace(value);
        }

        private static string Get(CsiTaskDto task, string key)
        {
            string value;
            return TryGetArgument(task, key, out value) ? value : null;
        }

        private static bool TryParseDouble(string value, out double result)
        {
            return double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out result);
        }

        private static double? ParseNullableDouble(string value)
        {
            double parsed;
            return TryParseDouble(value, out parsed) ? (double?)parsed : null;
        }

        private static List<string> SplitNames(string value)
        {
            var result = new List<string>();
            if (string.IsNullOrWhiteSpace(value))
            {
                return result;
            }

            string[] parts = value.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
            foreach (string part in parts)
            {
                string clean = CleanObjectName(part);
                if (!string.IsNullOrWhiteSpace(clean))
                {
                    AddUnique(result, clean);
                }
            }

            return result;
        }

        private static bool IsFrameCreation(CsiTaskDto task)
        {
            return string.Equals(task.TaskType, "FrameObj", StringComparison.OrdinalIgnoreCase) &&
                   string.Equals(task.Operation, "Add", StringComparison.OrdinalIgnoreCase);
        }

        private static void AddUnique(List<string> values, string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return;
            }

            foreach (string existing in values)
            {
                if (string.Equals(existing, value, StringComparison.OrdinalIgnoreCase))
                {
                    return;
                }
            }

            values.Add(value);
        }

        private static string CleanObjectName(string value)
        {
            return string.IsNullOrWhiteSpace(value) ? null : value.Trim().Trim('.', ',', ';', ':');
        }

        private static bool IsReservedName(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return true;
            }

            string normalized = value.Trim().ToLowerInvariant();
            return normalized == "point" ||
                   normalized == "frame" ||
                   normalized == "beam" ||
                   normalized == "section" ||
                   normalized == "load" ||
                   normalized == "then" ||
                   normalized == "and" ||
                   normalized == "to" ||
                   normalized == "from";
        }
    }
}
