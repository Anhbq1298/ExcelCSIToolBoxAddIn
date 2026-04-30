using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
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
            "points.get_selected_by_name",
            "points.get_guid",
            "points.get_group_assignments",
            "points.get_connectivity",
            "points.get_spring",
            "points.get_mass",
            "points.get_local_axes",
            "points.get_diaphragm",
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
            "execute_csi_request",
            "points.add_by_coordinates",
            "frames.add_by_coordinates",
            "frames.add_by_points",
            "frames.add_object",
            "frames.add_objects",
            "frames.assign_section",
            "loads.frame.assign_distributed",
            "loads.frame.assign_point_load",
            "random.generate_objects",
            "Workflow_CreateTruss",
            "truss.generate_howe",
            "truss.generate_pratt",
            "FrameObject_AssignDistributedLoad",
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
- points.get_guid, points.get_connectivity, points.get_group_assignments, points.get_spring, points.get_mass, points.get_local_axes, points.get_diaphragm: Point detail queries.
- frames.get_all_names, frames.get_sections, frames.get_section_detail, frames.get_count: Frame queries.
- shells.get_all_names, shells.get_selected, shells.get_property, shells.get_count: Shell queries.
- loads.combinations.get_all, loads.patterns.get_all: Loading queries.
- execute_csi_request: Multi-step CSI workflow tool. Use when the user asks for multiple actions in one request.
- points.add_by_coordinates, frames.add_object, frames.add_objects: Creation tools.
- random.generate_objects: Generate random CSI points, frames, and shell/area objects using safe defaults.
- truss.generate_howe: Generate a Howe truss with optional slope, chord/web sections, and distributed load assignment. Continuous chords; released vertical/brace members.
- truss.generate_pratt: Generate a Pratt truss with optional slope, chord/web sections, and distributed load assignment. Continuous chords; released vertical/brace members.

SAFETY POLICY:
1. Do not use dryRun unless the user explicitly asks for preview/check only.
2. For LOW risk (Add Point, Add Frame, Clear Selection), execute directly with dryRun: false and confirmed: true.
3. For HIGH/MEDIUM risk (Delete, Run Analysis, Save, Large Assignments), ask for explicit user confirmation before execution. When confirmed, call the tool with dryRun: false and confirmed: true.

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
        private readonly IAiToolDispatcher _toolDispatcher;
        private readonly AiAgentResponseBuilder _responseBuilder;
        private readonly CsiIntentPlannerService _intentPlannerService;
        private readonly AgentTaskPlannerService _taskPlannerService;
        private readonly AgentTaskExecutorService _taskExecutorService;
        private string _pendingToolName;
        private string _pendingArgumentsJson;

        public AiAgentOrchestrator(OllamaChatService ollamaChatService, LocalMcpClient mcpClient)
        {
            _ollamaChatService = ollamaChatService
                ?? throw new ArgumentNullException(nameof(ollamaChatService));
            _mcpClient = mcpClient
                ?? throw new ArgumentNullException(nameof(mcpClient));
            _toolDispatcher = new AiToolDispatcher(_mcpClient);
            _responseBuilder = new AiAgentResponseBuilder();
            _intentPlannerService = new CsiIntentPlannerService(_ollamaChatService);
            _taskPlannerService = new AgentTaskPlannerService();
            _taskExecutorService = new AgentTaskExecutorService(
                (taskText, token) => ExecuteSingleRequestAsync(taskText, null, token));
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

            IReadOnlyList<AgentTaskItem> plannedTasks = _taskPlannerService.CreateTasks(userMessage);
            if (plannedTasks.Count > 1)
            {
                string detectedText = AgentTaskExecutorService.FormatDetectedTasks(plannedTasks) + Environment.NewLine;
                onAssistantToken?.Invoke(detectedText);

                AgentTaskExecutionSummary taskSummary = await _taskExecutorService.ExecuteAsync(plannedTasks, cancellationToken);
                string outcomeText = AgentTaskExecutorService.FormatFinalResponse(taskSummary, false);
                onAssistantToken?.Invoke(outcomeText);
                bool anyToolCall = HasAnyToolCall(taskSummary);

                return new AiAgentResponse
                {
                    AssistantText = detectedText + outcomeText,
                    ToolWasCalled = anyToolCall,
                    RoutingReason = anyToolCall
                        ? "Request decomposition route: multiple tasks detected and executed in order."
                        : "Request decomposition route: multiple tasks detected; no MCP tool was called."
                };
            }

            return await ExecuteSingleRequestAsync(userMessage, onAssistantToken, cancellationToken);
        }

        private async Task<AiAgentResponse> ExecuteSingleRequestAsync(
            string userMessage,
            Action<string> onAssistantToken,
            CancellationToken cancellationToken)
        {
            AiAgentToolDecision decision = await _intentPlannerService.TryCreateToolDecisionAsync(userMessage, cancellationToken);
            if (decision != null && decision.ClarificationRequired)
            {
                string clarificationText = string.IsNullOrWhiteSpace(decision.ClarificationMessage)
                    ? "Please provide the missing action, target object, and required parameters."
                    : decision.ClarificationMessage;
                onAssistantToken?.Invoke(clarificationText);
                return _responseBuilder.Clarification(clarificationText, decision.Reason);
            }

            if (decision == null)
            {
                decision = TryCreateHeuristicToolDecision(userMessage);
            }

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
                string missingToolDiagnostic = TryCreateNoToolMatchedMessage(userMessage, decision);
                if (!string.IsNullOrWhiteSpace(missingToolDiagnostic))
                {
                    onAssistantToken?.Invoke(missingToolDiagnostic);
                    return new AiAgentResponse
                    {
                        AssistantText = missingToolDiagnostic,
                        ToolWasCalled = false,
                        RoutingReason = "No MCP tool matched the validated CSI request."
                    };
                }

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

            string decisionValidationFailure = TryRepairOrRejectInvalidFrameAddDecision(userMessage, ref decision);
            if (!string.IsNullOrWhiteSpace(decisionValidationFailure))
            {
                onAssistantToken?.Invoke(decisionValidationFailure);
                return new AiAgentResponse
                {
                    AssistantText = decisionValidationFailure,
                    ToolWasCalled = false,
                    RoutingReason = "Rejected invalid frame add tool arguments before MCP execution."
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

            ToolCallResponse toolResponse = await _toolDispatcher.DispatchAsync(decision, cancellationToken);

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

        private static bool HasAnyToolCall(AgentTaskExecutionSummary summary)
        {
            if (summary == null || summary.Tasks == null)
            {
                return false;
            }

            for (int i = 0; i < summary.Tasks.Count; i++)
            {
                AgentTaskItem task = summary.Tasks[i];
                if (task != null && task.ToolWasCalled)
                {
                    return true;
                }
            }

            return false;
        }

        private static string TryCreateNoToolMatchedMessage(string userMessage, AiAgentToolDecision decision)
        {
            if (decision != null && !string.IsNullOrWhiteSpace(decision.MissingSchemaMessage))
            {
                return decision.MissingSchemaMessage;
            }

            string expectedTool = TryInferExpectedToolName(userMessage);
            if (string.IsNullOrWhiteSpace(expectedTool))
            {
                return null;
            }

            return "No MCP tool matched. Missing tool schema: " + expectedTool + ".";
        }

        private static string TryInferExpectedToolName(string userMessage)
        {
            string normalized = Normalize(userMessage);
            if (ContainsAny(normalized, "truss", "howe", "pratt", "warren", "mono-slope", "monoslope"))
            {
                return "Workflow_CreateTruss";
            }

            if (ContainsAny(normalized, "udl", "distributed load", "uniform load", "top chord"))
            {
                return "FrameObject_AssignDistributedLoad";
            }

            return null;
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
                string operationName = result.Value<string>("OperationName");

                if (requiresConfirmation && !string.IsNullOrWhiteSpace(operationName))
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

            if (LooksLikeWorkflow(normalized))
            {
                JObject args = new JObject
                {
                    ["userInput"] = userMessage
                };

                return new AiAgentToolDecision
                {
                    ShouldCallTool = true,
                    ToolName = "execute_csi_request",
                    ArgumentsJson = args.ToString(Newtonsoft.Json.Formatting.None),
                    Reason = "Heuristic route: multi-step CSI workflow."
                };
            }

            AiAgentToolDecision addPointDecision = TryCreateAddPointDecision(userMessage);
            if (addPointDecision != null)
            {
                return addPointDecision;
            }

            AiAgentToolDecision addFrameDecision = TryCreateAddFrameDecision(userMessage);
            if (addFrameDecision != null)
            {
                return addFrameDecision;
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
                ContainsAny(normalized, "list", "count", "how many", "howmany", "number", "all", "names", "bao nhieu", "bao nhiêu", "dem", "đếm"))
            {
                return CreateToolDecision("loads.combinations.get_all", "Heuristic route: load combinations query.");
            }

            if (ContainsAny(normalized, "load pattern", "load patterns", "pattern", "patterns") &&
                ContainsAny(normalized, "list", "count", "how many", "howmany", "number", "all", "names", "bao nhieu", "bao nhiêu", "dem", "đếm"))
            {
                return CreateToolDecision("loads.patterns.get_all", "Heuristic route: load patterns query.");
            }

            if (ContainsAny(normalized, "selected point", "selected joint", "point selected", "joint selected"))
            {
                return CreateToolDecision("points.get_selected", "Heuristic route: selected points query.");
            }

            if (ContainsAny(normalized, "point", "points", "joint", "joints") &&
                ContainsAny(normalized, "count", "how many", "howmany", "number", "bao nhieu", "bao nhiêu", "dem", "đếm"))
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
                ContainsAny(normalized, "count", "how many", "howmany", "number", "bao nhieu", "bao nhiêu", "dem", "đếm"))
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
                ContainsAny(normalized, "how many", "howmany", "count", "number", "bao nhieu", "bao nhiêu", "dem", "đếm"))
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

        private static AiAgentToolDecision TryCreateAddPointDecision(string userMessage)
        {
            if (string.IsNullOrWhiteSpace(userMessage))
            {
                return null;
            }

            string normalized = Normalize(userMessage);
            if (!ContainsAny(normalized, "add point", "create point", "new point"))
            {
                return null;
            }

            Match coordinateMatch = Regex.Match(
                userMessage,
                @"(-?\d+(?:\.\d+)?)\s*,\s*(-?\d+(?:\.\d+)?)\s*,\s*(-?\d+(?:\.\d+)?)",
                RegexOptions.CultureInvariant);
            if (!coordinateMatch.Success)
            {
                return null;
            }

            double x;
            double y;
            double z;
            if (!double.TryParse(coordinateMatch.Groups[1].Value, NumberStyles.Float, CultureInfo.InvariantCulture, out x) ||
                !double.TryParse(coordinateMatch.Groups[2].Value, NumberStyles.Float, CultureInfo.InvariantCulture, out y) ||
                !double.TryParse(coordinateMatch.Groups[3].Value, NumberStyles.Float, CultureInfo.InvariantCulture, out z))
            {
                return null;
            }

            string pointName = string.Empty;
            Match nameMatch = Regex.Match(
                userMessage,
                @"(?:name\s+it\s+(?:as\s+)?|named\s+|name\s*[:=]\s*)([A-Za-z_][A-Za-z0-9_\-\.]*)",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
            if (nameMatch.Success)
            {
                pointName = nameMatch.Groups[1].Value;
            }

            JObject args = new JObject
            {
                ["x"] = x,
                ["y"] = y,
                ["z"] = z,
                ["userName"] = pointName,
                ["dryRun"] = false,
                ["confirmed"] = true
            };

            return new AiAgentToolDecision
            {
                ShouldCallTool = true,
                ToolName = "points.add_by_coordinates",
                ArgumentsJson = args.ToString(Newtonsoft.Json.Formatting.None),
                Reason = "Heuristic route: add point by coordinates."
            };
        }

        private static AiAgentToolDecision TryCreateAddFrameDecision(string userMessage)
        {
            if (string.IsNullOrWhiteSpace(userMessage))
            {
                return null;
            }

            string normalized = Normalize(userMessage);
            if (!ContainsAny(
                    normalized,
                    "add frame",
                    "create frame",
                    "new frame",
                    "draw frame",
                    "draw a frame",
                    "draw frame object",
                    "draw a frame object",
                    "add beam",
                    "create beam",
                    "new beam",
                    "draw beam",
                    "add member",
                    "create member",
                    "new member",
                    "draw member"))
            {
                return null;
            }

            AiAgentToolDecision coordinateDecision = TryCreateAddFrameByCoordinatesDecision(userMessage);
            if (coordinateDecision != null)
            {
                return coordinateDecision;
            }

            return TryCreateAddFrameByPointsDecision(userMessage);
        }

        private static AiAgentToolDecision TryCreateAddFrameByCoordinatesDecision(string userMessage)
        {
            MatchCollection coordinateMatches = Regex.Matches(
                userMessage,
                @"(-?\d+(?:\.\d+)?)\s*,\s*(-?\d+(?:\.\d+)?)\s*,\s*(-?\d+(?:\.\d+)?)",
                RegexOptions.CultureInvariant);
            if (coordinateMatches.Count < 2)
            {
                return null;
            }

            double xi;
            double yi;
            double zi;
            double xj;
            double yj;
            double zj;
            if (!TryParseCoordinateTriple(coordinateMatches[0], out xi, out yi, out zi) ||
                !TryParseCoordinateTriple(coordinateMatches[1], out xj, out yj, out zj))
            {
                return null;
            }

            JObject args = new JObject
            {
                ["xi"] = xi,
                ["yi"] = yi,
                ["zi"] = zi,
                ["xj"] = xj,
                ["yj"] = yj,
                ["zj"] = zj,
                ["propName"] = ExtractSectionName(userMessage),
                ["userName"] = ExtractObjectName(userMessage),
            };

            return new AiAgentToolDecision
            {
                ShouldCallTool = true,
                ToolName = "frames.add_object",
                ArgumentsJson = args.ToString(Newtonsoft.Json.Formatting.None),
                Reason = "Heuristic route: add frame by coordinates."
            };
        }

        private static AiAgentToolDecision TryCreateAddFrameByPointsDecision(string userMessage)
        {
            Match pointsMatch = Regex.Match(
                userMessage,
                @"(?:between|from)\s+(?:point\s+)?([A-Za-z_][A-Za-z0-9_\-\.]*)\s+(?:and|to)\s+(?:point\s+)?([A-Za-z_][A-Za-z0-9_\-\.]*)",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
            if (!pointsMatch.Success)
            {
                return null;
            }

            JObject args = new JObject
            {
                ["pointIName"] = pointsMatch.Groups[1].Value,
                ["pointJName"] = pointsMatch.Groups[2].Value,
                ["propName"] = ExtractSectionName(userMessage),
                ["userName"] = ExtractObjectName(userMessage),
            };

            return new AiAgentToolDecision
            {
                ShouldCallTool = true,
                ToolName = "frames.add_object",
                ArgumentsJson = args.ToString(Newtonsoft.Json.Formatting.None),
                Reason = "Heuristic route: add frame by points."
            };
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

        private static string TryRepairOrRejectInvalidFrameAddDecision(string userMessage, ref AiAgentToolDecision decision)
        {
            if (decision == null ||
                !string.Equals(decision.ToolName, "frames.add_object", StringComparison.OrdinalIgnoreCase))
            {
                return null;
            }

            JObject args;
            try
            {
                args = JObject.Parse(decision.ArgumentsJson ?? "{}");
            }
            catch
            {
                return "Failed to add frame: the tool arguments were not valid JSON.";
            }

            string pointI = ReadString(args, "PointIName", "pointIName");
            string pointJ = ReadString(args, "PointJName", "pointJName");
            if (string.IsNullOrWhiteSpace(pointI) && string.IsNullOrWhiteSpace(pointJ))
            {
                return null;
            }

            if (IsValidCsiPointName(pointI) && IsValidCsiPointName(pointJ))
            {
                return null;
            }

            AiAgentToolDecision coordinateDecision = TryCreateAddFrameByCoordinatesDecision(userMessage);
            if (coordinateDecision != null)
            {
                decision = coordinateDecision;
                return null;
            }

            return "Failed to add frame: the endpoints look like coordinates, but I could not find six valid coordinate values. Use this format: add frame from (0,0,0) to (10000,10000,1000).";
        }

        private static string Normalize(string text)
        {
            return (text ?? string.Empty).Trim().ToLowerInvariant();
        }

        private static string ReadString(JObject obj, params string[] names)
        {
            for (int i = 0; i < names.Length; i++)
            {
                JToken token = obj[names[i]];
                string value = token == null ? null : token.ToString();
                if (!string.IsNullOrWhiteSpace(value))
                {
                    return value;
                }
            }

            return null;
        }

        private static bool IsValidCsiPointName(string value)
        {
            return !string.IsNullOrWhiteSpace(value) &&
                   Regex.IsMatch(value.Trim(), @"^[A-Za-z_][A-Za-z0-9_\-\.]*$", RegexOptions.CultureInvariant);
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

        private static bool LooksLikeWorkflow(string normalized)
        {
            int verbCount = CountKeywordHits(normalized, "add", "create", "assign", "apply", "set", "get", "extract", "select", "update", "delete", "connect");
            int sequenceCount = CountKeywordHits(normalized, " then ", " after that ", " next ", " also ", " finally ", " followed by ");
            int objectTypeCount = CountKeywordHits(normalized, "point", "frame", "beam", "load", "section", "shell", "area", "combination", "case");

            return verbCount > 1 ||
                   sequenceCount > 0 ||
                   objectTypeCount > 1 && ContainsAny(normalized, " and ", " then ", " also ", ",");
        }

        private static int CountKeywordHits(string text, params string[] keywords)
        {
            int count = 0;
            for (int i = 0; i < keywords.Length; i++)
            {
                if (text.IndexOf(keywords[i], StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    count++;
                }
            }

            return count;
        }

        private static bool TryParseCoordinateTriple(Match match, out double x, out double y, out double z)
        {
            x = 0;
            y = 0;
            z = 0;

            return match != null &&
                match.Success &&
                double.TryParse(match.Groups[1].Value, NumberStyles.Float, CultureInfo.InvariantCulture, out x) &&
                double.TryParse(match.Groups[2].Value, NumberStyles.Float, CultureInfo.InvariantCulture, out y) &&
                double.TryParse(match.Groups[3].Value, NumberStyles.Float, CultureInfo.InvariantCulture, out z);
        }

        private static string ExtractObjectName(string userMessage)
        {
            Match nameMatch = Regex.Match(
                userMessage ?? string.Empty,
                @"(?:name\s+it\s+(?:as\s+)?|named\s+|name\s*[:=]\s*)([A-Za-z_][A-Za-z0-9_\-\.]*)",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
            return nameMatch.Success ? nameMatch.Groups[1].Value : string.Empty;
        }

        private static string ExtractSectionName(string userMessage)
        {
            Match sectionMatch = Regex.Match(
                userMessage ?? string.Empty,
                @"(?:section|property|prop)\s*(?:name)?\s*(?:as|=|:)?\s*([A-Za-z_][A-Za-z0-9_\-\.]*)",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
            return sectionMatch.Success ? sectionMatch.Groups[1].Value : "Default";
        }

        private static string TryFormatToolResponse(ToolCallResponse toolResponse)
        {
            if (toolResponse == null)
            {
                return null;
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
                    case "execute_csi_request":
                        return FormatCsiWorkflowResult(result);
                    case "frames.add_object":
                        return FormatFrameAddObjectResult(result);
                    case "frames.add_objects":
                        return FormatFrameAddObjectsResult(result);
                    case "random.generate_objects":
                        return FormatRandomGenerationResult(result);
                    case "truss.generate_howe":
                    case "truss.generate_pratt":
                        return FormatTrussResult(result);
                    default:
                        string preview = TryFormatWritePreview(result);
                        if (!string.IsNullOrWhiteSpace(preview))
                        {
                            return preview;
                        }

                        return null;
                }
            }
            catch
            {
                return toolResponse.Success ? null : toolResponse.Message;
            }
        }

        private static string FormatCsiWorkflowResult(JObject result)
        {
            int total = result.Value<int?>("TotalTasksDetected") ?? 0;
            int succeeded = result.Value<int?>("Succeeded") ?? 0;
            int failed = result.Value<int?>("Failed") ?? 0;
            int skipped = result.Value<int?>("Skipped") ?? 0;

            return failed == 0 && skipped == 0
                ? $"Task completed. {succeeded}/{total} workflow task(s) succeeded."
                : $"Task completed with issues. {succeeded}/{total} succeeded, {failed} failed, {skipped} skipped.";
        }

        private static string FormatRandomGenerationResult(JObject result)
        {
            int requestedPoints = result.Value<int?>("RequestedPoints") ?? 0;
            int requestedFrames = result.Value<int?>("RequestedFrames") ?? 0;
            int requestedShells = result.Value<int?>("RequestedShells") ?? 0;
            int addedPoints = result.Value<int?>("AddedPoints") ?? 0;
            int addedFrames = result.Value<int?>("AddedFrames") ?? 0;
            int addedShells = result.Value<int?>("AddedShells") ?? 0;
            int failedItems = result.Value<int?>("FailedItems") ?? 0;
            string summary = $"Task completed. Added {addedPoints}/{requestedPoints} point(s), {addedFrames}/{requestedFrames} frame(s), {addedShells}/{requestedShells} shell(s).";
            return failedItems == 0
                ? summary
                : summary + $" {failedItems} item(s) failed.";
        }

        private static string FormatTrussResult(JObject result)
        {
            bool success = result.Value<bool?>("Success") ?? false;
            int bayCount = result.Value<int?>("BayCount") ?? 0;
            int added = result.Value<int?>("AddedFrameCount") ?? 0;
            int released = result.Value<int?>("ReleasedWebMemberCount") ?? 0;
            int loaded = result.Value<int?>("LoadedFrameCount") ?? 0;
            double span = result.Value<double?>("Span") ?? 0;
            string trussType = result.Value<string>("TrussType") ?? "Howe";
            string slopeMode = result.Value<string>("SlopeMode");
            double slope = result.Value<double?>("Slope") ?? 0;
            string slopeText = slope > 0
                ? $", slope mode {slopeMode ?? "Gable"} ({slope.ToString("0.###", CultureInfo.InvariantCulture)})"
                : string.Empty;
            string loadText = loaded > 0 ? $" Assigned distributed load to {loaded} frame(s)." : string.Empty;
            string summary = $"Task completed. {trussType} truss generated with {bayCount} bay(s), span {span.ToString("0.###", CultureInfo.InvariantCulture)}{slopeText}. Added {added} frame(s); released {released} web member(s).{loadText}";
            return success ? summary : "Task completed with issues. " + summary;
        }

        private static string FormatFrameAddObjectResult(JObject result)
        {
            bool success = result.Value<bool?>("Success") ?? false;
            string frameName = result.Value<string>("FrameName");
            string addMethod = result.Value<string>("AddMethod") ?? "FrameObj";
            string failureReason = result.Value<string>("FailureReason");
            int? returnCode = result.Value<int?>("ReturnCode");

            if (success)
            {
                return string.IsNullOrWhiteSpace(frameName)
                    ? "Task completed. Added 1 frame."
                    : $"Task completed. Added frame {frameName}.";
            }

            string reason = string.IsNullOrWhiteSpace(failureReason)
                ? returnCode.HasValue ? $"CSI return code {returnCode.Value}." : "The frame was not added."
                : failureReason;

            return $"Task failed. Frame was not added: {reason}";
        }

        private static string FormatFrameAddObjectsResult(JObject result)
        {
            int total = result.Value<int?>("TotalRequested") ?? 0;
            int successCount = result.Value<int?>("SuccessCount") ?? 0;
            int failureCount = result.Value<int?>("FailureCount") ?? 0;
            return failureCount == 0
                ? $"Task completed. Added {successCount}/{total} frame(s)."
                : $"Task completed with issues. Added {successCount}/{total} frame(s); {failureCount} failed.";
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

        private static string TryFormatWritePreview(JObject result)
        {
            string operationName = result.Value<string>("OperationName");
            if (string.IsNullOrWhiteSpace(operationName))
            {
                return null;
            }

            string summary = result.Value<string>("Summary");
            bool requiresConfirmation = result.Value<bool?>("RequiresConfirmation") ?? false;

            if (requiresConfirmation)
            {
                return "Preview: " + (summary ?? "This operation requires confirmation.") + " Confirm to proceed?";
            }

            return "Preview only: " + (summary ?? "No model changes were made.");
        }
    }
}

