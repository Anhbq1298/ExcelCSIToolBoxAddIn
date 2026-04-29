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

            AiAgentToolDecision trussDecision = TryCreateHoweTrussDecision(userMessage);
            if (trussDecision != null)
            {
                return trussDecision;
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

        private static AiAgentToolDecision TryCreateHoweTrussDecision(string userMessage)
        {
            if (string.IsNullOrWhiteSpace(userMessage) ||
                !Regex.IsMatch(userMessage, @"\b(howe|pratt|truss)\b", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant))
            {
                return null;
            }

            bool isPratt = Regex.IsMatch(userMessage, @"\bpratt\b", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
            string trussType = isPratt ? "Pratt" : "Howe";
            JObject args = new JObject
            {
                ["TrussType"] = trussType
            };
            int bayCount = ExtractBayCount(userMessage);
            if (bayCount > 0)
            {
                args["BayCount"] = bayCount;
            }

            double span = ExtractDimension(userMessage, @"\b(?:span|length)\s*(?:=|:|is|of)?\s*(?<value>\d+(?:\.\d+)?)");
            if (span > 0)
            {
                args["Span"] = span;
            }

            double height = ExtractDimension(userMessage, @"\bheight\s*(?:=|:|is|of)?\s*(?<value>\d+(?:\.\d+)?)");
            if (height > 0)
            {
                args["Height"] = height;
            }

            double slope = ExtractSlope(userMessage);
            if (slope > 0)
            {
                args["Slope"] = slope;
                args["SlopeMode"] = ExtractSlopeMode(userMessage);
                args["MonoSlopeDirection"] = ExtractMonoSlopeDirection(userMessage);
            }

            string prefix = ExtractName(userMessage, @"\b(?:prefix|name)\s*(?:=|:|is|as)?\s*(?<name>[A-Za-z_][A-Za-z0-9_\-\.]*)");
            if (!string.IsNullOrWhiteSpace(prefix))
            {
                args["NamePrefix"] = prefix;
            }

            string chordSection = ExtractName(userMessage, @"\b(?:chord|chords|top\s+chord|bottom\s+chord)\s+(?:section|property|prop)\s*(?:=|:|is|as)?\s*(?<name>[A-Za-z_][A-Za-z0-9_\-\.]*)");
            if (!string.IsNullOrWhiteSpace(chordSection))
            {
                args["ChordPropName"] = chordSection;
            }

            string webSection = ExtractName(userMessage, @"\b(?:web|webs|brace|braces|vertical|verticals)\s+(?:section|property|prop)\s*(?:=|:|is|as)?\s*(?<name>[A-Za-z_][A-Za-z0-9_\-\.]*)");
            if (!string.IsNullOrWhiteSpace(webSection))
            {
                args["WebPropName"] = webSection;
            }

            string section = ExtractName(userMessage, @"\b(?:section|property|prop)\s*(?:=|:|is|as)?\s*(?<name>[A-Za-z_][A-Za-z0-9_\-\.]*)");
            if (!string.IsNullOrWhiteSpace(section))
            {
                if (string.IsNullOrWhiteSpace(chordSection))
                {
                    args["ChordPropName"] = section;
                }

                if (string.IsNullOrWhiteSpace(webSection))
                {
                    args["WebPropName"] = section;
                }
            }

            double distributedLoad = ExtractDistributedLoadValue(userMessage);
            if (distributedLoad != 0)
            {
                args["DistributedLoadPattern"] = ExtractLoadPattern(userMessage) ?? "DEAD";
                args["DistributedLoadDirection"] = ExtractLoadDirection(userMessage);
                args["DistributedLoadValue1"] = distributedLoad;
                args["DistributedLoadValue2"] = distributedLoad;
                args["DistributedLoadTarget"] = ExtractDistributedLoadTarget(userMessage);
            }

            return new AiAgentToolDecision
            {
                ShouldCallTool = true,
                ToolName = isPratt ? "truss.generate_pratt" : "truss.generate_howe",
                ArgumentsJson = args.ToString(Formatting.None),
                Reason = $"Intent planner deterministic route: {trussType} truss generation."
            };
        }

        private static int ExtractBayCount(string text)
        {
            Match match = Regex.Match(
                text ?? string.Empty,
                @"\b(?<count>\d{1,3})\s*(?:bay|bays)\b|\b(?:bay|bays)\s*(?:=|:)?\s*(?<count2>\d{1,3})\b",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
            if (!match.Success)
            {
                return 0;
            }

            string valueText = match.Groups["count"].Success ? match.Groups["count"].Value : match.Groups["count2"].Value;
            int value;
            return int.TryParse(valueText, NumberStyles.Integer, CultureInfo.InvariantCulture, out value) ? value : 0;
        }

        private static double ExtractDimension(string text, string pattern)
        {
            Match match = Regex.Match(text ?? string.Empty, pattern, RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
            if (!match.Success)
            {
                return 0;
            }

            double value;
            return double.TryParse(match.Groups["value"].Value, NumberStyles.Float, CultureInfo.InvariantCulture, out value) ? value : 0;
        }

        private static string ExtractName(string text, string pattern)
        {
            Match match = Regex.Match(text ?? string.Empty, pattern, RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
            return match.Success ? match.Groups["name"].Value : null;
        }

        private static double ExtractDistributedLoadValue(string text)
        {
            string source = text ?? string.Empty;
            if (!Regex.IsMatch(source, @"\b(?:load|udl|distributed)\b", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant))
            {
                return 0;
            }

            Match afterKeyword = Regex.Match(
                source,
                @"\b(?:udl|distributed\s+load|load)\s*(?:=|:|is|of)?\s*(?<value>-?\d+(?:\.\d+)?)",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
            if (afterKeyword.Success && !IsLikelyLoadPatternValue(source, afterKeyword.Index))
            {
                double value;
                if (double.TryParse(afterKeyword.Groups["value"].Value, NumberStyles.Float, CultureInfo.InvariantCulture, out value))
                {
                    return value;
                }
            }

            Match beforeKeyword = Regex.Match(
                source,
                @"(?<value>-?\d+(?:\.\d+)?)\s*(?:kn/m|n/m|kip/ft|plf|k/ft)?\s*(?:udl|distributed\s+load|load)\b",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
            if (beforeKeyword.Success)
            {
                double value;
                if (double.TryParse(beforeKeyword.Groups["value"].Value, NumberStyles.Float, CultureInfo.InvariantCulture, out value))
                {
                    return value;
                }
            }

            return 0;
        }

        private static bool IsLikelyLoadPatternValue(string text, int matchIndex)
        {
            int start = Math.Max(0, matchIndex - 12);
            string prefix = text.Substring(start, matchIndex - start);
            return Regex.IsMatch(prefix, @"pattern\s*$", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
        }

        private static string ExtractLoadPattern(string text)
        {
            string pattern = ExtractName(text, @"\b(?:load\s+pattern|pattern|loadpat)\s*(?:=|:|is|as)?\s*(?<name>[A-Za-z_][A-Za-z0-9_\-\.]*)");
            if (!string.IsNullOrWhiteSpace(pattern))
            {
                return pattern;
            }

            if (Regex.IsMatch(text ?? string.Empty, @"\bdead\b", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant))
            {
                return "DEAD";
            }

            if (Regex.IsMatch(text ?? string.Empty, @"\blive\b", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant))
            {
                return "LIVE";
            }

            return null;
        }

        private static int ExtractLoadDirection(string text)
        {
            Match direction = Regex.Match(
                text ?? string.Empty,
                @"\b(?:dir|direction)\s*(?:=|:|is)?\s*(?<value>\d{1,2})\b",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
            if (direction.Success)
            {
                int value;
                if (int.TryParse(direction.Groups["value"].Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out value) &&
                    value > 0)
                {
                    return value;
                }
            }

            return 6;
        }

        private static string ExtractDistributedLoadTarget(string text)
        {
            string source = text ?? string.Empty;
            if (Regex.IsMatch(source, @"\b(?:load|udl|distributed\s+load)\b[^.;]*(?:to|on|onto|for)?\s*(?:all\s+members|all\s+frames|entire\s+truss)\b", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant))
            {
                return "All";
            }

            if (Regex.IsMatch(source, @"\b(?:web|webs|brace|braces|vertical|verticals)\s+(?:load|udl)\b|\b(?:load|udl|distributed\s+load)\b[^.;]*(?:to|on|onto|for)\s+(?:web|webs|brace|braces|vertical|verticals)\b", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant))
            {
                return "Web";
            }

            if (Regex.IsMatch(source, @"\b(?:bottom\s+chord|bottom)\s+(?:load|udl)\b|\b(?:load|udl|distributed\s+load)\b[^.;]*(?:to|on|onto|for)\s+(?:bottom\s+chord|bottom)\b", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant))
            {
                return "BottomChord";
            }

            if (Regex.IsMatch(source, @"\b(?:chords|both\s+chords)\s+(?:load|udl)\b|\b(?:load|udl|distributed\s+load)\b[^.;]*(?:to|on|onto|for)\s+(?:chords|both\s+chords)\b", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant))
            {
                return "Chord";
            }

            return "TopChord";
        }

        private static double ExtractSlope(string text)
        {
            string source = text ?? string.Empty;

            Match ratio = Regex.Match(
                source,
                @"\b(?:slope|pitch)\s*(?:=|:|is|of)?\s*(?<rise>\d+(?:\.\d+)?)\s*[:/]\s*(?<run>\d+(?:\.\d+)?)",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
            if (ratio.Success)
            {
                double rise;
                double run;
                if (double.TryParse(ratio.Groups["rise"].Value, NumberStyles.Float, CultureInfo.InvariantCulture, out rise) &&
                    double.TryParse(ratio.Groups["run"].Value, NumberStyles.Float, CultureInfo.InvariantCulture, out run) &&
                    run > 0)
                {
                    return rise / run;
                }
            }

            Match percent = Regex.Match(
                source,
                @"\b(?:slope|pitch)\s*(?:=|:|is|of)?\s*(?<value>\d+(?:\.\d+)?)\s*%",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
            if (percent.Success)
            {
                double value;
                return double.TryParse(percent.Groups["value"].Value, NumberStyles.Float, CultureInfo.InvariantCulture, out value)
                    ? value / 100.0
                    : 0;
            }

            Match degree = Regex.Match(
                source,
                @"\b(?:slope|pitch)\s*(?:=|:|is|of)?\s*(?<value>\d+(?:\.\d+)?)\s*(?:deg|degree|degrees)\b",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
            if (degree.Success)
            {
                double value;
                return double.TryParse(degree.Groups["value"].Value, NumberStyles.Float, CultureInfo.InvariantCulture, out value)
                    ? Math.Tan(value * Math.PI / 180.0)
                    : 0;
            }

            return ExtractDimension(source, @"\b(?:slope|pitch)\s*(?:=|:|is|of)?\s*(?<value>\d+(?:\.\d+)?)");
        }

        private static string ExtractSlopeMode(string text)
        {
            string source = text ?? string.Empty;
            if (Regex.IsMatch(
                source,
                @"\b(?:mono\s*slope|monoslope|one\s*side|single\s*slope|from\s+(?:one|1)\s+side)\b",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant))
            {
                return "Mono";
            }

            if (Regex.IsMatch(
                source,
                @"\b(?:middle|center|centre|both\s+sides?|two\s+sides?|gable|double\s*slope)\b",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant))
            {
                return "Gable";
            }

            return "Gable";
        }

        private static string ExtractMonoSlopeDirection(string text)
        {
            string source = text ?? string.Empty;
            if (Regex.IsMatch(
                source,
                @"\b(?:from\s+right|right\s+to\s+left|high\s+(?:at|on)\s+left|low\s+(?:at|on)\s+right)\b",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant))
            {
                return "RightToLeft";
            }

            return "LeftToRight";
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
                @"\b(csi|etabs|sap2000|model|unit|point|joint|frame|beam|member|column|brace|shell|area|slab|wall|panel|section|property|load|udl|select|selection|length|random|truss|howe|pratt)\b|-?\d+(?:\.\d+)?\s*,\s*-?\d+(?:\.\d+)?\s*,\s*-?\d+(?:\.\d+)?",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
        }
    }
}
