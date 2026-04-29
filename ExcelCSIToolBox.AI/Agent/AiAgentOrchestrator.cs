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
            "points.get_selected",
            "points.get_all_names",
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
            "frames.get_by_name",
            "frames.get_points",
            "frames.get_section",
            "frames.get_distributed_loads",
            "frames.get_point_loads",
            "frames.set_selected",
            "frames.get_sections",
            "frames.get_section_detail",
            "shells.get_all_names",
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
@"You are a safe local tool router for an Excel CSI toolbox.
Decide whether the user's message requires a local MCP tool.

Available read tools:
1. CSI.GetModelInfo
   Use when the user asks about current model file, model path, model name, or attached model info.

2. CSI.GetPresentUnits
   Use when the user asks about current model units.

3. CSI.GetSelectedObjects
   Use when the user asks about selected objects or current selection.

4. CSI.GetSelectedFrames
   Use when the user asks about selected frames.

5. CSI.GetSelectedFrameSections
   Use when the user asks about section properties of selected frames.

6. shells.get_all_names
   Use when the user asks how many shell, area, wall, slab, membrane, plate, or shell elements/objects exist in the running model.

7. shells.get_by_name
   Args: {""areaName"":""A1""}
   Use when the user asks for details of a named shell/area object.

8. shells.get_points
   Args: {""areaName"":""A1""}
   Use when the user asks which points define a shell/area object.

9. shells.get_property
   Args: {""areaName"":""A1""}
   Use when the user asks for the assigned shell/area property.

10. shells.get_selected
   Use when the user asks about selected shell/area/wall/slab objects.

11. shells.get_uniform_loads
   Args: {""areaName"":""A1""}
   Use when the user asks for uniform loads assigned to a shell/area object.

12. points.get_selected
   Use when the user asks about selected points/joints or their coordinates.

13. points.get_all_names
   Use when the user asks to list/count all point/joint objects.

14. points.get_by_name
   Args: {""pointName"":""P1""}
   Use when the user asks for details of a named point/joint object.

15. points.get_coordinates
   Args: {""pointName"":""P1""}
   Use when the user asks for coordinates of a named point/joint object.

16. frames.get_selected
   Use when the user asks about selected frames/beams/columns.

17. frames.get_all_names
   Use when the user asks to list/count all frame/beam/column/brace objects.

18. frames.get_by_name
   Args: {""frameName"":""F1""}
   Use when the user asks for details of a named frame object.

19. frames.get_points
   Args: {""frameName"":""F1""}
   Use when the user asks for end points of a named frame object.

20. frames.get_section
   Args: {""frameName"":""F1""}
   Use when the user asks for section assignment of a named frame object.

21. frames.get_sections
   Use when the user asks for available frame section properties.

22. frames.get_section_detail
   Args: {""sectionName"":""W18X35""}
   Use when the user asks for dimensions or material of a frame section.

23. loads.combinations.get_all
   Use when the user asks to list/count load combinations.

24. loads.combinations.get_details
   Args: {""combinationName"":""COMB1""}
   Use when the user asks for cases/scale factors inside one load combination.

25. loads.patterns.get_all
   Use when the user asks to list/count load patterns.

Available controlled write tools:
26. points.add_by_coordinates
   Args: {""dryRun"":true,""confirmed"":false,""x"":0,""y"":0,""z"":0,""userName"":""""}
20. frames.add_by_coordinates
   Args: {""dryRun"":true,""confirmed"":false,""xi"":0,""yi"":0,""zi"":0,""xj"":0,""yj"":0,""zj"":0,""sectionName"":"""",""userName"":""""}
21. frames.add_by_points
   Args: {""dryRun"":true,""confirmed"":false,""point1Name"":"""",""point2Name"":"""",""sectionName"":"""",""userName"":""""}
22. frames.assign_section
   Args: {""dryRun"":true,""confirmed"":false,""frameNames"":[""F1""],""sectionName"":""""}
23. loads.frame.assign_distributed
   Args: {""dryRun"":true,""confirmed"":false,""frameNames"":[""F1""],""loadPattern"":"""",""direction"":6,""value1"":0,""value2"":0}
24. loads.frame.assign_point_load
   Args: {""dryRun"":true,""confirmed"":false,""frameNames"":[""F1""],""loadPattern"":"""",""direction"":6,""distance"":0.5,""value"":0}
25. selection.clear
   Args: {""dryRun"":true,""confirmed"":false}
26. frames.delete
   Args: {""dryRun"":true,""confirmed"":false,""objectNames"":[""F1""]}
27. analysis.run
   Args: {""dryRun"":true,""confirmed"":false}
28. file.save_model
   Dangerous and blocked by default.
29. shells.add_by_points
   Args: {""dryRun"":true,""confirmed"":false,""pointNames"":[""P1"",""P2"",""P3""],""propertyName"":""Default"",""userName"":""""}
30. shells.add_by_coordinates
   Args: {""dryRun"":true,""confirmed"":false,""points"":[{""x"":0,""y"":0,""z"":0},{""x"":1,""y"":0,""z"":0},{""x"":0,""y"":1,""z"":0}],""propertyName"":""Default"",""userName"":"""",""coordinateSystem"":""Global""}
31. shells.assign_uniform_load
   Args: {""dryRun"":true,""confirmed"":false,""areaNames"":[""A1""],""loadPattern"":""DEAD"",""value"":1.0,""direction"":6,""replace"":true,""coordinateSystem"":""Global""}
32. shells.delete
   Args: {""dryRun"":true,""confirmed"":false,""areaNames"":[""A1""]}
33. points.set_selected
   Args: {""dryRun"":true,""confirmed"":false,""names"":[""P1"",""P2""]}
34. frames.set_selected
   Args: {""dryRun"":true,""confirmed"":false,""names"":[""F1"",""F2""]}
35. loads.combinations.delete
   Args: {""dryRun"":true,""confirmed"":false,""names"":[""COMB1""]}
36. loads.patterns.delete
   Args: {""dryRun"":true,""confirmed"":false,""names"":[""DEAD""]}
37. points.set_restraint
   Args: {""dryRun"":true,""confirmed"":false,""pointNames"":[""P1""],""restraints"":[true,true,true,false,false,false]}
38. points.set_load_force
   Args: {""dryRun"":true,""confirmed"":false,""pointNames"":[""P1""],""loadPattern"":""DEAD"",""forceValues"":[0,0,-10,0,0,0],""replace"":true,""coordinateSystem"":""Global""}
39. frames.assign_distributed_load
   Args: {""dryRun"":true,""confirmed"":false,""frameNames"":[""F1""],""loadPattern"":""DEAD"",""direction"":6,""value1"":0,""value2"":0}
40. frames.assign_point_load
   Args: {""dryRun"":true,""confirmed"":false,""frameNames"":[""F1""],""loadPattern"":""DEAD"",""direction"":6,""distance"":0.5,""value"":0}

Return JSON only:
{
  ""shouldCallTool"": false,
  ""toolName"": """",
  ""argumentsJson"": ""{}"",
  ""reason"": """"
}

Strict rules:
- Only choose tools from the available tool list.
- Never invent a tool.
- For any model modification, set dryRun=true and confirmed=false first.
- Never set confirmed=true unless the user is clearly confirming a preview.
- Refuse unlock, open model, raw COM calls, mass delete, and unlisted tools.
- Do not invent live model data.";

        private const string NormalChatSystemPrompt =
@"You are an AI assistant embedded inside an Excel CSI toolbox.
You can help the user understand Excel, ETABS, SAP2000, C# code, and structural engineering workflows.
You may query and modify CSI models only through approved local MCP tools.
For model modification requests, always run a dry-run preview first, explain affected objects, and wait for explicit confirmation.
Never use raw COM calls or unlisted tools.
Keep responses concise and practical.";

        private const string ToolResultSummarySystemPrompt =
@"You are an AI assistant inside an Excel CSI toolbox.
The user asked a question.
A local MCP tool was called and returned structured JSON.
Summarize the result clearly for the user.
If this is a write preview requiring confirmation, ask the user to confirm before execution.
If a write operation was blocked, explain why simply.
If the result indicates failure, explain the failure simply and suggest checking whether
ETABS/SAP2000 is open and a model is loaded.";

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
                return confirmationResponse;
            }

            AiAgentToolDecision decision = TryCreateHeuristicToolDecision(userMessage);
            if (decision == null)
            {
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
                string chatText = await _ollamaChatService.ChatAsync(
                    new List<OllamaMessage>
                    {
                        new OllamaMessage { role = "system", content = NormalChatSystemPrompt },
                        new OllamaMessage { role = "user", content = userMessage }
                    },
                    cancellationToken);

                return new AiAgentResponse
                {
                    AssistantText = chatText,
                    ToolWasCalled = false,
                    RoutingReason = decision.Reason
                };
            }

            if (!ApprovedTools.Contains(decision.ToolName))
            {
                string refusalText = await _ollamaChatService.ChatAsync(
                    new List<OllamaMessage>
                    {
                        new OllamaMessage { role = "system", content = NormalChatSystemPrompt },
                        new OllamaMessage { role = "user", content = userMessage }
                    },
                    cancellationToken);

                return new AiAgentResponse
                {
                    AssistantText = "That tool is not approved for safe local execution. " + refusalText,
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

            string summaryText = await _ollamaChatService.ChatAsync(
                new List<OllamaMessage>
                {
                    new OllamaMessage { role = "system", content = ToolResultSummarySystemPrompt },
                    new OllamaMessage { role = "user", content = toolResultContext }
                },
                cancellationToken);

            return new AiAgentResponse
            {
                AssistantText = summaryText,
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
                ContainsAny(normalized, "list", "count", "how many", "number", "all", "names", "bao nhieu", "bao nhiêu", "dem", "đếm"))
            {
                return CreateToolDecision("points.get_all_names", "Heuristic route: point object query.");
            }

            if (ContainsAny(normalized, "selected frame", "selected member", "frame selected", "selected beam", "selected column"))
            {
                return CreateToolDecision("frames.get_selected", "Heuristic route: selected frames query.");
            }

            if (ContainsAny(normalized, "frame", "frames", "member", "members", "beam", "beams", "column", "columns", "brace", "braces") &&
                ContainsAny(normalized, "list", "count", "how many", "number", "all", "names", "bao nhieu", "bao nhiêu", "dem", "đếm"))
            {
                return CreateToolDecision("frames.get_all_names", "Heuristic route: frame object query.");
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
                ContainsAny(normalized, "how many", "count", "number", "list", "names", "bao nhieu", "bao nhiêu", "dem", "đếm"))
            {
                return CreateToolDecision("shells.get_all_names", "Heuristic route: shell/area object query.");
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
    }
}
