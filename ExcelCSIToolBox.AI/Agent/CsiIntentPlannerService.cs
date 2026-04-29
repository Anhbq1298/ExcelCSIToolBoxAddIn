using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using ExcelCSIToolBox.AI.Ollama;
using ExcelCSIToolBox.Data.CSISapModel.Intent;
using ExcelCSIToolBox.Data.CSISapModel.Workflow;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace ExcelCSIToolBox.AI.Agent
{
    public sealed class CsiIntentPlannerService
    {
        private const string PlannerSystemPrompt =
@"You convert natural language CSI/ETABS/SAP2000 requests into canonical executable tasks.
Return JSON only. No markdown. No explanation.

Schema:
{
  ""tasks"": [
    {
      ""taskType"": ""PointObj | FrameObj | FrameLoad | Selection | Query | Model"",
      ""operation"": ""AddCartesian | Add | AssignSection | AssignDistributed | SelectFrames | ExtractFrameLengths | GetInfo | GetPresentUnits"",
      ""arguments"": { ""key"": ""value"" },
      ""dependsOn"": []
    }
  ]
}

Rules:
- Use PointObj/AddCartesian for point creation. Arguments: name, x, y, z.
- Use FrameObj/Add for frame/beam/member creation.
- FrameObj/Add by points arguments: userName, pointIName, pointJName, propName.
- FrameObj/Add by coordinates arguments: userName, xi, yi, zi, xj, yj, zj, propName.
- Synonyms such as draw, model, place, insert, create, add, connect can mean Add.
- If the user gives two coordinate triples like 0,0,0 to 6900,6900,6900 for a frame, use coordinate arguments.
- Never use pointIName or pointJName for values that look like coordinates, numbers, or parenthesized coordinate text.
- Use FrameObj/AssignSection for section/property assignment. Arguments: frameNames, sectionName.
- Use FrameLoad/AssignDistributed for UDL/distributed load. Arguments: frameNames, loadPattern, direction, value1, value2.
- Use Query/ExtractFrameLengths for frame length requests. Arguments: frameNames.
- Use Model/GetPresentUnits for units requests.
- Use Model/GetInfo for model info/path/file requests.
- For multi-step requests, return tasks in execution order.
- If a task depends on an earlier created object, leave dependsOn empty; the executor will infer basic dependencies.
- If the request is not a CSI model task, return { ""tasks"": [] }.
- Do not invent missing coordinates or point names.";

        private readonly OllamaChatService _ollamaChatService;

        public CsiIntentPlannerService(OllamaChatService ollamaChatService)
        {
            _ollamaChatService = ollamaChatService ?? throw new ArgumentNullException(nameof(ollamaChatService));
        }

        public async Task<AiAgentToolDecision> TryCreateToolDecisionAsync(
            string userMessage,
            CancellationToken cancellationToken)
        {
            if (!ShouldUsePlanner(userMessage))
            {
                return null;
            }

            AiAgentToolDecision randomDecision = TryCreateRandomGenerationDecision(userMessage);
            if (randomDecision != null)
            {
                return randomDecision;
            }

            AiAgentToolDecision deterministicDecision = TryCreateDeterministicFrameCoordinateDecision(userMessage);
            if (deterministicDecision != null)
            {
                return deterministicDecision;
            }

            CsiIntentPlanDto plan = await TryCreatePlanAsync(userMessage, cancellationToken);
            if (plan == null || plan.Tasks == null || plan.Tasks.Count == 0)
            {
                return null;
            }

            List<CsiTaskDto> tasks = NormalizeTasks(plan.Tasks);
            if (tasks.Count == 0)
            {
                return null;
            }

            AiAgentToolDecision directDecision = TryCreateDirectToolDecision(tasks);
            if (directDecision != null)
            {
                return directDecision;
            }

            JObject args = new JObject
            {
                ["userInput"] = userMessage,
                ["plannedTasks"] = JArray.FromObject(tasks)
            };

            return new AiAgentToolDecision
            {
                ShouldCallTool = true,
                ToolName = "execute_csi_request",
                ArgumentsJson = args.ToString(Formatting.None),
                Reason = "Intent planner route: canonical CSI workflow."
            };
        }

        private async Task<CsiIntentPlanDto> TryCreatePlanAsync(
            string userMessage,
            CancellationToken cancellationToken)
        {
            try
            {
                string response = await _ollamaChatService.ChatAsync(
                    new List<OllamaMessage>
                    {
                        new OllamaMessage { role = "system", content = PlannerSystemPrompt },
                        new OllamaMessage { role = "user", content = userMessage }
                    },
                    cancellationToken);

                string json = ExtractJsonObject(response);
                if (string.IsNullOrWhiteSpace(json))
                {
                    return null;
                }

                return JsonConvert.DeserializeObject<CsiIntentPlanDto>(json);
            }
            catch
            {
                return null;
            }
        }

        private static List<CsiTaskDto> NormalizeTasks(IReadOnlyList<CsiIntentTaskDto> intentTasks)
        {
            var tasks = new List<CsiTaskDto>();
            for (int i = 0; i < intentTasks.Count; i++)
            {
                CsiIntentTaskDto intentTask = intentTasks[i];
                if (intentTask == null ||
                    string.IsNullOrWhiteSpace(intentTask.TaskType) ||
                    string.IsNullOrWhiteSpace(intentTask.Operation))
                {
                    continue;
                }

                var task = new CsiTaskDto
                {
                    TaskId = "task-" + (tasks.Count + 1).ToString(CultureInfo.InvariantCulture),
                    TaskType = intentTask.TaskType.Trim(),
                    Operation = intentTask.Operation.Trim(),
                    Arguments = NormalizeArguments(intentTask.Arguments),
                    DependsOn = intentTask.DependsOn ?? new List<string>()
                };

                if (IsValidTask(task))
                {
                    tasks.Add(task);
                }
            }

            return tasks;
        }

        private static Dictionary<string, string> NormalizeArguments(Dictionary<string, string> arguments)
        {
            var normalized = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            if (arguments == null)
            {
                return normalized;
            }

            foreach (KeyValuePair<string, string> item in arguments)
            {
                if (string.IsNullOrWhiteSpace(item.Key))
                {
                    continue;
                }

                normalized[item.Key.Trim()] = item.Value == null ? null : item.Value.Trim();
            }

            return normalized;
        }

        private static bool IsValidTask(CsiTaskDto task)
        {
            if (IsTask(task, "PointObj", "AddCartesian"))
            {
                return HasNumber(task, "x") && HasNumber(task, "y") && HasNumber(task, "z");
            }

            if (IsTask(task, "FrameObj", "Add"))
            {
                return HasValidPointName(task, "pointIName") && HasValidPointName(task, "pointJName") ||
                       HasNumber(task, "xi") && HasNumber(task, "yi") && HasNumber(task, "zi") &&
                       HasNumber(task, "xj") && HasNumber(task, "yj") && HasNumber(task, "zj");
            }

            if (IsTask(task, "FrameObj", "AssignSection"))
            {
                return HasText(task, "sectionName");
            }

            if (IsTask(task, "FrameLoad", "AssignDistributed"))
            {
                return HasText(task, "loadPattern") &&
                       HasNumber(task, "value1") &&
                       HasNumber(task, "value2");
            }

            if (IsTask(task, "Selection", "SelectFrames") ||
                IsTask(task, "Query", "ExtractFrameLengths") ||
                IsTask(task, "Model", "GetInfo") ||
                IsTask(task, "Model", "GetPresentUnits"))
            {
                return true;
            }

            return false;
        }

        private static AiAgentToolDecision TryCreateDirectToolDecision(IReadOnlyList<CsiTaskDto> tasks)
        {
            if (tasks.Count != 1)
            {
                return null;
            }

            CsiTaskDto task = tasks[0];
            if (IsTask(task, "FrameObj", "Add"))
            {
                JObject args = new JObject();
                Copy(args, task, "userName", "UserName");
                Copy(args, task, "propName", "PropName");
                Copy(args, task, "pointIName", "PointIName");
                Copy(args, task, "pointJName", "PointJName");
                Copy(args, task, "xi", "Xi");
                Copy(args, task, "yi", "Yi");
                Copy(args, task, "zi", "Zi");
                Copy(args, task, "xj", "Xj");
                Copy(args, task, "yj", "Yj");
                Copy(args, task, "zj", "Zj");

                return new AiAgentToolDecision
                {
                    ShouldCallTool = true,
                    ToolName = "frames.add_object",
                    ArgumentsJson = args.ToString(Formatting.None),
                    Reason = "Intent planner route: add frame."
                };
            }

            if (IsTask(task, "PointObj", "AddCartesian"))
            {
                JObject args = new JObject
                {
                    ["X"] = Get(task, "x"),
                    ["Y"] = Get(task, "y"),
                    ["Z"] = Get(task, "z"),
                    ["UserName"] = Get(task, "name"),
                    ["dryRun"] = false,
                    ["confirmed"] = true
                };

                return new AiAgentToolDecision
                {
                    ShouldCallTool = true,
                    ToolName = "points.add_by_coordinates",
                    ArgumentsJson = args.ToString(Formatting.None),
                    Reason = "Intent planner route: add point."
                };
            }

            if (IsTask(task, "Model", "GetPresentUnits"))
            {
                return CreateDecision("CSI.GetPresentUnits", "Intent planner route: units query.");
            }

            if (IsTask(task, "Model", "GetInfo"))
            {
                return CreateDecision("CSI.GetModelInfo", "Intent planner route: model info query.");
            }

            return null;
        }

        private static AiAgentToolDecision CreateDecision(string toolName, string reason)
        {
            return new AiAgentToolDecision
            {
                ShouldCallTool = true,
                ToolName = toolName,
                ArgumentsJson = "{}",
                Reason = reason
            };
        }

        private static void Copy(JObject args, CsiTaskDto task, string sourceName, string targetName)
        {
            string value = Get(task, sourceName);
            if (!string.IsNullOrWhiteSpace(value))
            {
                args[targetName] = value;
            }
        }

        private static bool IsTask(CsiTaskDto task, string taskType, string operation)
        {
            return string.Equals(task.TaskType, taskType, StringComparison.OrdinalIgnoreCase) &&
                   string.Equals(task.Operation, operation, StringComparison.OrdinalIgnoreCase);
        }

        private static bool HasText(CsiTaskDto task, string key)
        {
            return !string.IsNullOrWhiteSpace(Get(task, key));
        }

        private static bool HasValidPointName(CsiTaskDto task, string key)
        {
            string value = Get(task, key);
            return !string.IsNullOrWhiteSpace(value) &&
                   Regex.IsMatch(value, @"^[A-Za-z_][A-Za-z0-9_\-\.]*$", RegexOptions.CultureInvariant);
        }

        private static bool HasNumber(CsiTaskDto task, string key)
        {
            double value;
            return double.TryParse(Get(task, key), NumberStyles.Float, CultureInfo.InvariantCulture, out value);
        }

        private static string Get(CsiTaskDto task, string key)
        {
            string value;
            return task.Arguments != null && task.Arguments.TryGetValue(key, out value) ? value : null;
        }

        private static string ExtractJsonObject(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return null;
            }

            Match fenced = Regex.Match(text, "```(?:json)?\\s*(?<json>\\{[\\s\\S]*?\\})\\s*```", RegexOptions.IgnoreCase);
            if (fenced.Success)
            {
                return fenced.Groups["json"].Value;
            }

            int start = text.IndexOf('{');
            int end = text.LastIndexOf('}');
            return start >= 0 && end > start ? text.Substring(start, end - start + 1) : null;
        }

        private static AiAgentToolDecision TryCreateDeterministicFrameCoordinateDecision(string userMessage)
        {
            if (string.IsNullOrWhiteSpace(userMessage) ||
                !Regex.IsMatch(userMessage, @"\b(frame|beam|member)\b", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant))
            {
                return null;
            }

            MatchCollection coordinateMatches = Regex.Matches(
                userMessage,
                @"\(?\s*(-?\d+(?:\.\d+)?)\s*,\s*(-?\d+(?:\.\d+)?)\s*,\s*(-?\d+(?:\.\d+)?)\s*\)?",
                RegexOptions.CultureInvariant);
            if (coordinateMatches.Count < 2)
            {
                return null;
            }

            JObject args = new JObject
            {
                ["Xi"] = coordinateMatches[0].Groups[1].Value,
                ["Yi"] = coordinateMatches[0].Groups[2].Value,
                ["Zi"] = coordinateMatches[0].Groups[3].Value,
                ["Xj"] = coordinateMatches[1].Groups[1].Value,
                ["Yj"] = coordinateMatches[1].Groups[2].Value,
                ["Zj"] = coordinateMatches[1].Groups[3].Value
            };

            return new AiAgentToolDecision
            {
                ShouldCallTool = true,
                ToolName = "frames.add_object",
                ArgumentsJson = args.ToString(Formatting.None),
                Reason = "Intent planner deterministic route: add frame by coordinate triples."
            };
        }

        private static AiAgentToolDecision TryCreateRandomGenerationDecision(string userMessage)
        {
            if (string.IsNullOrWhiteSpace(userMessage) ||
                !Regex.IsMatch(userMessage, @"\brandom\b", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant))
            {
                return null;
            }

            bool points = Regex.IsMatch(userMessage, @"\b(point|points|joint|joints)\b", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
            bool frames = Regex.IsMatch(userMessage, @"\b(frame|frames|beam|beams|member|members|column|columns|brace|braces)\b", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
            bool shells = Regex.IsMatch(userMessage, @"\b(shell|shells|area|areas|slab|slabs|wall|walls|panel|panels)\b", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
            if (!points && !frames && !shells)
            {
                return null;
            }

            JObject args = new JObject
            {
                ["AddPoints"] = points,
                ["AddFrames"] = frames,
                ["AddShells"] = shells
            };

            int count = ExtractFirstInteger(userMessage);
            if (count > 0)
            {
                if (points)
                {
                    args["PointCount"] = count;
                }

                if (frames)
                {
                    args["FrameCount"] = count;
                }

                if (shells)
                {
                    args["ShellCount"] = count;
                }
            }

            return new AiAgentToolDecision
            {
                ShouldCallTool = true,
                ToolName = "random.generate_objects",
                ArgumentsJson = args.ToString(Formatting.None),
                Reason = "Intent planner deterministic route: random CSI object generation."
            };
        }

        private static int ExtractFirstInteger(string text)
        {
            Match match = Regex.Match(text ?? string.Empty, @"\b(?<count>\d{1,3})\b", RegexOptions.CultureInvariant);
            if (!match.Success)
            {
                return 0;
            }

            int value;
            return int.TryParse(match.Groups["count"].Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out value)
                ? value
                : 0;
        }

        private static bool ShouldUsePlanner(string userMessage)
        {
            if (string.IsNullOrWhiteSpace(userMessage))
            {
                return false;
            }

            return Regex.IsMatch(
                userMessage,
                @"\b(csi|etabs|sap2000|model|unit|point|joint|frame|beam|member|column|brace|shell|area|slab|wall|panel|section|property|load|udl|select|selection|length|random)\b|-?\d+(?:\.\d+)?\s*,\s*-?\d+(?:\.\d+)?\s*,\s*-?\d+(?:\.\d+)?",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
        }
    }
}
