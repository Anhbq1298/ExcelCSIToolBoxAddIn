using System;
using System.Collections.Generic;
using System.Globalization;
using System.Threading;
using System.Threading.Tasks;
using ExcelCSIToolBox.AI.Mcp.Client;
using ExcelCSIToolBox.AI.Mcp.Contracts;
using ExcelCSIToolBox.AI.Ollama;
using Newtonsoft.Json.Linq;

namespace ExcelCSIToolBox.AI.Agent
{
    /// <summary>
    /// Orchestrates the full AI agent conversation loop.
    /// Uses fast heuristic routing for common CSI queries to avoid unnecessary LLM passes.
    /// </summary>
    public class AiAgentOrchestrator
    {
        private static readonly HashSet<string> ApprovedTools = new HashSet<string>(
            StringComparer.OrdinalIgnoreCase)
        {
            "CSI.GetModelInfo",
            "CSI.GetPresentUnits",
            "CSI.GetSelectedObjects",
            "CSI.GetSelectedFrames",
            "CSI.GetSelectedFrameSections",
            "csi.get_model_statistics",
            "csi.refresh_view",
            "points.get_selected",
            "points.get_all_names",
            "points.get_count",
            "points.get_by_name",
            "points.get_coordinates",
            "points.get_restraint",
            "points.get_load_forces",
            "points.set_selected",
            "points.add_cartesian",
            "points.set_restraint",
            "points.set_load_force",
            "frames.get_selected",
            "frames.get_all_names",
            "frames.get_count",
            "frames.get_by_name",
            "frames.get_points",
            "frames.get_section",
            "frames.get_distributed_loads",
            "frames.get_point_loads",
            "frames.set_selected",
            "frames.get_sections",
            "frames.get_section_detail",
            "frames.analyze_selected",
            "shells.get_all_names",
            "shells.get_count",
            "shells.get_by_name",
            "shells.get_points",
            "shells.get_property",
            "shells.get_selected",
            "shells.get_uniform_loads",
            "shells.add_by_points",
            "shells.add_by_coordinates",
            "shells.assign_uniform_load",
            "shells.delete",
            "loads.combinations.get_all",
            "loads.combinations.get_details",
            "loads.combinations.delete",
            "loads.patterns.get_all",
            "loads.patterns.delete",
            "points.add_by_coordinates",
            "frames.add_by_coordinates",
            "frames.add_by_points",
            "frames.assign_section",
            "loads.frame.assign_distributed",
            "loads.frame.assign_point_load",
            "frames.assign_distributed_load",
            "frames.assign_point_load",
            "selection.clear",
            "frames.delete",
            "analysis.run",
            "file.save_model"
        };

        private const string ToolDecisionSystemPrompt =
@"You are an assistant inside an Excel CSI toolbox.
Prefer MCP tools over guessing model data.
Be concise and engineering-focused.
Treat all model operations as read-only unless the user explicitly requests a write action.

Available tools:
- CSI.GetModelInfo: File path, product info.
- csi.get_model_statistics: Counts for points, frames, shells, loads. Use for ""how many"" questions.
- CSI.GetPresentUnits: Current model units.
- CSI.GetSelectedObjects: List current selection.
- csi.refresh_view: Refresh graphics.
- frames.analyze_selected: COMPLETE workflow for selected frames (sections, geometry, assignments).
- points.get_all_names, points.get_coordinates, points.get_selected, points.get_count: Point queries.
- frames.get_all_names, frames.get_sections, frames.get_section_detail, frames.get_count: Frame queries.
- shells.get_all_names, shells.get_selected, shells.get_property, shells.get_count: Shell queries.
- loads.combinations.get_all, loads.patterns.get_all: Loading queries.
- points.add_by_coordinates, frames.add_by_coordinates, frames.add_by_points: Creation tools.

SAFETY POLICY:
1. For HIGH/MEDIUM risk (Delete, Run Analysis, Save, Large Assignments), you MUST run a dry-run preview first.
2. For LOW risk (Add Point, Add Frame, Clear Selection), you may execute directly (dryRun: false, confirmed: true) to be fast, unless the user specifically asks to preview first.

Return JSON only:
{
  ""shouldCallTool"": false,
  ""toolName"": """",
  ""argumentsJson"": ""{}"",
  ""reason"": """"
}";

        private const string NormalChatSystemPrompt =
@"You are an assistant inside an Excel CSI toolbox. 
Keep responses concise and structural engineering focused. 
Use MCP tools for any model-related queries.";

        private const string ToolResultSummarySystemPrompt =
@"Summarize the tool result clearly. 
If this is a write preview, ask for explicit confirmation before execution.";

        private readonly OllamaChatService _ollamaChatService;
        private readonly LocalMcpClient _mcpClient;
        private string _pendingToolName;
        private string _pendingArgumentsJson;

        public AiAgentOrchestrator(OllamaChatService ollamaChatService, LocalMcpClient mcpClient)
        {
            _ollamaChatService = ollamaChatService
                ?? throw new ArgumentNullException(nameof(ollamaChatService));
            _mcpClient = mcpClient
                ?? throw new ArgumentNullException(nameof(mcpClient));
        }

        public async Task<AiAgentResponse> SendAsync(
            string userMessage,
            Action<string> onAssistantToken,
            CancellationToken cancellationToken)
        {
            if (string.IsNullOrWhiteSpace(userMessage))
            {
                return new AiAgentResponse
                {
                    AssistantText = "Please enter a message.",
                    ToolWasCalled = false
                };
            }

            AiAgentResponse confirmationResponse = await TryHandlePendingConfirmationAsync(userMessage, cancellationToken);
            if (confirmationResponse != null)
            {
                onAssistantToken?.Invoke(confirmationResponse.AssistantText);
                return confirmationResponse;
            }

            AiAgentToolDecision decision = TryCreateHeuristicToolDecision(userMessage);
            if (decision == null)
            {
                // We usually don't stream the tool decision as it should be fast and JSON-only.
                string decisionResponse = await _ollamaChatService.ChatAsync(
                    new List<OllamaMessage>
                    {
                        new OllamaMessage { role = "system", content = ToolDecisionSystemPrompt },
                        new OllamaMessage { role = "user", content = userMessage }
                    },
                    cancellationToken);

                decision = AiAgentToolDecisionParser.Parse(decisionResponse);
            }

            if (!decision.ShouldCallTool || string.IsNullOrWhiteSpace(decision.ToolName))
            {
                var fullText = new System.Text.StringBuilder();
                await _ollamaChatService.ChatStreamAsync(
                    new List<OllamaMessage>
                    {
                        new OllamaMessage { role = "system", content = NormalChatSystemPrompt },
                        new OllamaMessage { role = "user", content = userMessage }
                    },
                    token =>
                    {
                        fullText.Append(token);
                        onAssistantToken?.Invoke(token);
                    },
                    cancellationToken);

                return new AiAgentResponse
                {
                    AssistantText = fullText.ToString(),
                    ToolWasCalled = false,
                    RoutingReason = decision.Reason
                };
            }

            if (!ApprovedTools.Contains(decision.ToolName) && !decision.ToolName.Contains("analyze_selected"))
            {
                string refusalText = "That tool is not approved for safe local execution.";
                onAssistantToken?.Invoke(refusalText);
                return new AiAgentResponse
                {
                    AssistantText = refusalText,
                    ToolWasCalled = false,
                    RoutingReason = $"Rejected non-approved tool: {decision.ToolName}"
                };
            }

            ToolCallResponse toolResponse = await _mcpClient.CallToolAsync(
                decision.ToolName,
                decision.ArgumentsJson,
                cancellationToken);

            TryRememberPendingWritePreview(toolResponse, decision.ArgumentsJson);

            string fastToolResponse = TryFormatToolResponse(toolResponse);
            if (!string.IsNullOrWhiteSpace(fastToolResponse))
            {
                onAssistantToken?.Invoke(fastToolResponse);
                return new AiAgentResponse
                {
                    AssistantText = fastToolResponse,
                    ToolWasCalled = true,
                    ToolName = toolResponse.ToolName,
                    ToolArgumentsJson = decision.ArgumentsJson,
                    ToolResponse = toolResponse,
                    RoutingReason = decision.Reason
                };
            }

            string toolResultContext =
                $"User question: {userMessage}\n\n" +
                $"Tool called: {toolResponse.ToolName}\n" +
                $"Tool success: {toolResponse.Success}\n" +
                $"Tool message: {toolResponse.Message}\n" +
                $"Tool result JSON:\n{toolResponse.ResultJson ?? "(none)"}";

            var summaryBuilder = new System.Text.StringBuilder();
            await _ollamaChatService.ChatStreamAsync(
                new List<OllamaMessage>
                {
                    new OllamaMessage { role = "system", content = ToolResultSummarySystemPrompt },
                    new OllamaMessage { role = "user", content = toolResultContext }
                },
                token =>
                {
                    summaryBuilder.Append(token);
                    onAssistantToken?.Invoke(token);
                },
                cancellationToken);

            return new AiAgentResponse
            {
                AssistantText = summaryBuilder.ToString(),
                ToolWasCalled = true,
                ToolName = toolResponse.ToolName,
                ToolArgumentsJson = decision.ArgumentsJson,
                ToolResponse = toolResponse,
                RoutingReason = decision.Reason
            };
        }

        private async Task<AiAgentResponse> TryHandlePendingConfirmationAsync(
            string userMessage,
            CancellationToken cancellationToken)
        {
            if (string.IsNullOrWhiteSpace(_pendingToolName) ||
                string.IsNullOrWhiteSpace(_pendingArgumentsJson))
            {
                return null;
            }

            string normalized = Normalize(userMessage);
            if (ContainsAny(normalized, "cancel", "no", "khong", "không", "huy", "hủy", "stop"))
            {
                _pendingToolName = null;
                _pendingArgumentsJson = null;
                return new AiAgentResponse
                {
                    AssistantText = "Cancelled. I did not execute the pending model change.",
                    ToolWasCalled = false
                };
            }

            if (!ContainsAny(normalized, "confirm", "yes", "ok", "proceed", "execute", "do it", "lam di", "làm đi", "dong y", "đồng ý"))
            {
                return null;
            }

            string toolName = _pendingToolName;
            string executeArguments = BuildConfirmedArguments(_pendingArgumentsJson);
            _pendingToolName = null;
            _pendingArgumentsJson = null;

            ToolCallResponse toolResponse = await _mcpClient.CallToolAsync(
                toolName,
                executeArguments,
                cancellationToken);

            return new AiAgentResponse
            {
                AssistantText = toolResponse.Success
                    ? toolResponse.Message
                    : "I could not execute the confirmed operation. " + toolResponse.Message,
                ToolWasCalled = true,
                ToolName = toolResponse.ToolName,
                ToolArgumentsJson = executeArguments,
                ToolResponse = toolResponse,
                RoutingReason = "Executed previously previewed write tool after user confirmation."
            };
        }

        private void TryRememberPendingWritePreview(ToolCallResponse toolResponse, string argumentsJson)
        {
            if (toolResponse == null ||
                !toolResponse.Success ||
                string.IsNullOrWhiteSpace(toolResponse.ResultJson))
            {
                return;
            }

            try
            {
                JObject result = JObject.Parse(toolResponse.ResultJson);
                bool requiresConfirmation = result.Value<bool?>("RequiresConfirmation") ?? false;
                bool supportsDryRun = result.Value<bool?>("SupportsDryRun") ?? false;
                string operationName = result.Value<string>("OperationName");

                if ((requiresConfirmation || supportsDryRun) && !string.IsNullOrWhiteSpace(operationName))
                {
                    _pendingToolName = operationName;
                    _pendingArgumentsJson = argumentsJson ?? "{}";
                }
            }
            catch
            {
                // Tool result was not a write preview; no pending confirmation needed.
            }
        }

        private static string BuildConfirmedArguments(string argumentsJson)
        {
            JObject args;
            try
            {
                args = JObject.Parse(argumentsJson ?? "{}");
            }
            catch
            {
                args = new JObject();
            }

            args["dryRun"] = false;
            args["confirmed"] = true;
            return args.ToString(Newtonsoft.Json.Formatting.None);
        }

        private static AiAgentToolDecision TryCreateHeuristicToolDecision(string userMessage)
        {
            string normalized = Normalize(userMessage);
            if (string.IsNullOrWhiteSpace(normalized))
            {
                return null;
            }

            if (ContainsAny(normalized, "unit", "units", "present unit", "current unit", "don vi"))
            {
                return CreateToolDecision("CSI.GetPresentUnits", "Heuristic route: units query.");
            }

            if (ContainsAny(normalized, "section") &&
                ContainsAny(normalized, "selected frame", "selected member", "frame selected", "selected beam", "selected column"))
            {
                return CreateToolDecision("CSI.GetSelectedFrameSections", "Heuristic route: selected frame sections query.");
            }

            if (ContainsAny(normalized, "load combination", "load combinations", "combo", "combination") &&
                ContainsAny(normalized, "list", "count", "how many", "number", "all", "names", "bao nhieu", "bao nhiêu", "dem", "đếm"))
            {
                return CreateToolDecision("loads.combinations.get_all", "Heuristic route: load combinations query.");
            }

            if (ContainsAny(normalized, "load pattern", "load patterns", "pattern", "patterns") &&
                ContainsAny(normalized, "list", "count", "how many", "number", "all", "names", "bao nhieu", "bao nhiêu", "dem", "đếm"))
            {
                return CreateToolDecision("loads.patterns.get_all", "Heuristic route: load patterns query.");
            }

            if (ContainsAny(normalized, "selected point", "selected joint", "point selected", "joint selected"))
            {
                return CreateToolDecision("points.get_selected", "Heuristic route: selected points query.");
            }

            if (ContainsAny(normalized, "point", "points", "joint", "joints") &&
                ContainsAny(normalized, "count", "how many", "number", "bao nhieu", "bao nhiêu", "dem", "đếm"))
            {
                return CreateToolDecision("points.get_count", "Heuristic route: point count query.");
            }

            if (ContainsAny(normalized, "point", "points", "joint", "joints") &&
                ContainsAny(normalized, "list", "all", "names"))
            {
                return CreateToolDecision("points.get_all_names", "Heuristic route: point object list.");
            }

            if (ContainsAny(normalized, "selected frame", "selected member", "frame selected", "selected beam", "selected column"))
            {
                return CreateToolDecision("frames.get_selected", "Heuristic route: selected frames query.");
            }

            if (ContainsAny(normalized, "frame", "frames", "member", "members", "beam", "beams", "column", "columns", "brace", "braces") &&
                ContainsAny(normalized, "count", "how many", "number", "bao nhieu", "bao nhiêu", "dem", "đếm"))
            {
                return CreateToolDecision("frames.get_count", "Heuristic route: frame count query.");
            }

            if (ContainsAny(normalized, "frame", "frames", "member", "members", "beam", "beams", "column", "columns", "brace", "braces") &&
                ContainsAny(normalized, "list", "all", "names"))
            {
                return CreateToolDecision("frames.get_all_names", "Heuristic route: frame object list.");
            }

            if (ContainsAny(normalized, "selected object", "current selection", "objects selected"))
            {
                return CreateToolDecision("CSI.GetSelectedObjects", "Heuristic route: current selection query.");
            }

            if (ContainsAny(normalized, "selected shell", "selected shells", "selected area", "selected areas", "selected wall", "selected slab"))
            {
                return CreateToolDecision("shells.get_selected", "Heuristic route: selected shell/area query.");
            }

            if (ContainsAny(normalized, "shell", "shells", "area", "areas", "wall", "walls", "slab", "slabs", "plate", "plates", "membrane", "membranes") &&
                ContainsAny(normalized, "how many", "count", "number", "bao nhieu", "bao nhiêu", "dem", "đếm"))
            {
                return CreateToolDecision("shells.get_count", "Heuristic route: shell count query.");
            }

            if (ContainsAny(normalized, "shell", "shells", "area", "areas", "wall", "walls", "slab", "slabs", "plate", "plates", "membrane", "membranes") &&
                ContainsAny(normalized, "list", "all", "names"))
            {
                return CreateToolDecision("shells.get_all_names", "Heuristic route: shell object list.");
            }

            if (ContainsAny(normalized, "model info", "model path", "model file", "file path", "attached model", "current model"))
            {
                return CreateToolDecision("CSI.GetModelInfo", "Heuristic route: model info query.");
            }

            return null;
        }

        private static AiAgentToolDecision CreateToolDecision(string toolName, string reason)
        {
            return new AiAgentToolDecision
            {
                ShouldCallTool = true,
                ToolName = toolName,
                ArgumentsJson = "{}",
                Reason = reason
            };
        }

        private static string Normalize(string text)
        {
            return (text ?? string.Empty).Trim().ToLowerInvariant();
        }

        private static bool ContainsAny(string text, params string[] values)
        {
            for (int i = 0; i < values.Length; i++)
            {
                if (text.IndexOf(values[i], StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return true;
                }
            }

            return false;
        }

        private static string TryFormatToolResponse(ToolCallResponse toolResponse)
        {
            if (toolResponse == null)
            {
                return null;
            }

            if (!toolResponse.Success)
            {
                return toolResponse.Message;
            }

            if (string.IsNullOrWhiteSpace(toolResponse.ResultJson))
            {
                return toolResponse.Message;
            }

            try
            {
                JObject result = JObject.Parse(toolResponse.ResultJson);

                switch (toolResponse.ToolName)
                {
                    case "CSI.GetPresentUnits":
                        return FormatPresentUnits(result);
                    case "CSI.GetModelInfo":
                        return FormatModelInfo(result);
                    case "CSI.GetSelectedFrames":
                        return FormatSelectedFrames(result);
                    case "CSI.GetSelectedObjects":
                        return FormatSelectedObjects(result);
                    case "CSI.GetSelectedFrameSections":
                        return FormatSelectedFrameSections(result);
                    case "points.get_selected":
                        return FormatSelectedPoints(result);
                    case "points.get_all_names":
                        return FormatNames(result, "point");
                    case "frames.get_selected":
                        return FormatNames(result, "selected frame");
                    case "frames.get_all_names":
                        return FormatNames(result, "frame");
                    case "loads.combinations.get_all":
                        return FormatLoadCombinations(result);
                    case "loads.patterns.get_all":
                        return FormatLoadPatterns(result);
                    case "shells.get_all_names":
                        return FormatShellNames(result);
                    case "shells.get_selected":
                        return FormatSelectedShells(result);
                    case "csi.get_model_statistics":
                        return FormatModelStatistics(result);
                    case "points.get_count":
                        return FormatCount(result, "point");
                    case "frames.get_count":
                        return FormatCount(result, "frame");
                    case "shells.get_count":
                        return FormatCount(result, "shell/area");
                    default:
                        return null;
                }
            }
            catch
            {
                return null;
            }
        }

        private static string FormatPresentUnits(JObject result)
        {
            string units = result.Value<string>("Units");
            return string.IsNullOrWhiteSpace(units)
                ? "Present units are unavailable."
                : "Current model units: " + units;
        }

        private static string FormatModelInfo(JObject result)
        {
            string product = result.Value<string>("Product") ?? "Unknown";
            string modelFile = result.Value<string>("ModelFile") ?? "Unknown";
            string modelPath = result.Value<string>("ModelPath") ?? "(unsaved model)";
            string currentUnit = result.Value<string>("CurrentUnit") ?? "Units unavailable";

            return $"Connected to {product}. Model file: {modelFile}. Path: {modelPath}. Units: {currentUnit}.";
        }

        private static string FormatSelectedFrames(JObject result)
        {
            int count = result.Value<int?>("Count") ?? 0;
            JArray frameNames = result["FrameNames"] as JArray;
            string preview = JoinPreview(frameNames, 10);

            if (count <= 0)
            {
                return "No frame objects are currently selected.";
            }

            return string.IsNullOrWhiteSpace(preview)
                ? $"Found {count.ToString(CultureInfo.InvariantCulture)} selected frame(s)."
                : $"Found {count.ToString(CultureInfo.InvariantCulture)} selected frame(s): {preview}.";
        }

        private static string FormatSelectedObjects(JObject result)
        {
            int count = result.Value<int?>("Count") ?? 0;
            JArray objects = result["Objects"] as JArray;
            if (count <= 0 || objects == null || objects.Count == 0)
            {
                return "No objects are currently selected.";
            }

            var preview = new List<string>();
            for (int i = 0; i < objects.Count && i < 8; i++)
            {
                JObject item = objects[i] as JObject;
                if (item == null)
                {
                    continue;
                }

                string objectType = item.Value<string>("ObjectType") ?? "Object";
                string uniqueName = item.Value<string>("UniqueName") ?? "?";
                preview.Add(objectType + " " + uniqueName);
            }

            string summary = string.Join(", ", preview);
            if (objects.Count > preview.Count)
            {
                summary += ", ...";
            }

            return $"Found {count.ToString(CultureInfo.InvariantCulture)} selected object(s): {summary}.";
        }

        private static string FormatSelectedFrameSections(JObject result)
        {
            int count = result.Value<int?>("Count") ?? 0;
            JArray assignments = result["Assignments"] as JArray;
            if (count <= 0 || assignments == null || assignments.Count == 0)
            {
                return "No selected frame sections were found.";
            }

            var preview = new List<string>();
            for (int i = 0; i < assignments.Count && i < 8; i++)
            {
                JObject item = assignments[i] as JObject;
                if (item == null)
                {
                    continue;
                }

                string frameName = item.Value<string>("FrameName") ?? "?";
                string sectionName = item.Value<string>("SectionName") ?? "(unavailable)";
                preview.Add(frameName + ": " + sectionName);
            }

            string summary = string.Join(", ", preview);
            if (assignments.Count > preview.Count)
            {
                summary += ", ...";
            }

            return $"Retrieved sections for {count.ToString(CultureInfo.InvariantCulture)} selected frame(s): {summary}.";
        }

        private static string FormatShellNames(JObject result)
        {
            string product = result.Value<string>("Product") ?? "the attached model";
            int count = result.Value<int?>("Count") ?? 0;
            JArray shellNames = result["ShellNames"] as JArray;
            string preview = JoinPreview(shellNames, 10);

            if (count <= 0)
            {
                return $"No shell/area objects were found in {product}.";
            }

            return string.IsNullOrWhiteSpace(preview)
                ? $"Found {count.ToString(CultureInfo.InvariantCulture)} shell/area object(s) in {product}."
                : $"Found {count.ToString(CultureInfo.InvariantCulture)} shell/area object(s) in {product}: {preview}.";
        }

        private static string FormatSelectedPoints(JObject result)
        {
            JArray points = result["Data"] as JArray;
            if (points == null || points.Count == 0)
            {
                return "No point objects are currently selected.";
            }

            var preview = new List<string>();
            for (int i = 0; i < points.Count && i < 8; i++)
            {
                JObject item = points[i] as JObject;
                if (item == null)
                {
                    continue;
                }

                string name = item.Value<string>("PointUniqueName") ?? item.Value<string>("PointLabel") ?? "?";
                double x = item.Value<double?>("X") ?? 0;
                double y = item.Value<double?>("Y") ?? 0;
                double z = item.Value<double?>("Z") ?? 0;
                preview.Add($"{name} ({x.ToString(CultureInfo.InvariantCulture)}, {y.ToString(CultureInfo.InvariantCulture)}, {z.ToString(CultureInfo.InvariantCulture)})");
            }

            string summary = string.Join(", ", preview);
            if (points.Count > preview.Count)
            {
                summary += ", ...";
            }

            return $"Found {points.Count.ToString(CultureInfo.InvariantCulture)} selected point(s): {summary}.";
        }

        private static string FormatNames(JObject result, string itemLabel)
        {
            JArray names = result["Data"] as JArray;
            if (names == null || names.Count == 0)
            {
                return $"No {itemLabel} objects were found.";
            }

            string preview = JoinPreview(names, 10);
            return string.IsNullOrWhiteSpace(preview)
                ? $"Found {names.Count.ToString(CultureInfo.InvariantCulture)} {itemLabel}(s)."
                : $"Found {names.Count.ToString(CultureInfo.InvariantCulture)} {itemLabel}(s): {preview}.";
        }

        private static string FormatLoadCombinations(JObject result)
        {
            JArray combinations = result["Data"] as JArray;
            if (combinations == null || combinations.Count == 0)
            {
                return "No load combinations were found.";
            }

            var preview = new List<string>();
            for (int i = 0; i < combinations.Count && i < 10; i++)
            {
                JObject item = combinations[i] as JObject;
                if (item == null)
                {
                    continue;
                }

                string name = item.Value<string>("Name") ?? "?";
                string type = item.Value<string>("Type") ?? "Unknown";
                preview.Add(name + " (" + type + ")");
            }

            string summary = string.Join(", ", preview);
            if (combinations.Count > preview.Count)
            {
                summary += ", ...";
            }

            return $"Found {combinations.Count.ToString(CultureInfo.InvariantCulture)} load combination(s): {summary}.";
        }

        private static string FormatLoadPatterns(JObject result)
        {
            JArray patterns = result["Data"] as JArray;
            if (patterns == null || patterns.Count == 0)
            {
                return "No load patterns were found.";
            }

            var preview = new List<string>();
            for (int i = 0; i < patterns.Count && i < 10; i++)
            {
                JObject item = patterns[i] as JObject;
                if (item == null)
                {
                    continue;
                }

                string name = item.Value<string>("Name") ?? "?";
                string type = item.Value<string>("Type") ?? "Unknown";
                preview.Add(name + " (" + type + ")");
            }

            string summary = string.Join(", ", preview);
            if (patterns.Count > preview.Count)
            {
                summary += ", ...";
            }

            return $"Found {patterns.Count.ToString(CultureInfo.InvariantCulture)} load pattern(s): {summary}.";
        }

        private static string FormatSelectedShells(JObject result)
        {
            JArray names = result["Data"] as JArray;
            if (names == null)
            {
                names = result["ShellNames"] as JArray;
            }

            int count = names == null ? 0 : names.Count;
            string preview = JoinPreview(names, 10);

            if (count <= 0)
            {
                return "No shell/area objects are currently selected.";
            }

            return string.IsNullOrWhiteSpace(preview)
                ? $"Found {count.ToString(CultureInfo.InvariantCulture)} selected shell/area object(s)."
                : $"Found {count.ToString(CultureInfo.InvariantCulture)} selected shell/area object(s): {preview}.";
        }

        private static string JoinPreview(JArray items, int maxItems)
        {
            if (items == null || items.Count == 0)
            {
                return string.Empty;
            }

            var preview = new List<string>();
            for (int i = 0; i < items.Count && i < maxItems; i++)
            {
                string value = items[i]?.ToString();
                if (!string.IsNullOrWhiteSpace(value))
                {
                    preview.Add(value);
                }
            }

            string summary = string.Join(", ", preview);
            if (items.Count > preview.Count)
            {
                summary += ", ...";
            }

            return summary;
        }

        private static string FormatModelStatistics(JObject result)
        {
            int points = result.Value<int?>("PointCount") ?? 0;
            int frames = result.Value<int?>("FrameCount") ?? 0;
            int shells = result.Value<int?>("ShellCount") ?? 0;
            int patterns = result.Value<int?>("LoadPatternCount") ?? 0;
            int combos = result.Value<int?>("LoadCombinationCount") ?? 0;

            return $"Model Statistics: {points} point(s), {frames} frame(s), {shells} shell/area(s), {patterns} load pattern(s), {combos} load combination(s).";
        }

        private static string FormatCount(JObject result, string label)
        {
            int count = result.Value<int?>("Count") ?? 0;
            return $"There are {count} {label} object(s) in the active model.";
        }
    }
}

