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
        private static readonly HashSet<string> ApprovedReadOnlyTools = new HashSet<string>(
            StringComparer.OrdinalIgnoreCase)
        {
            "CSI.GetModelInfo",
            "CSI.GetPresentUnits",
            "CSI.GetSelectedObjects",
            "CSI.GetSelectedFrames",
            "CSI.GetSelectedFrameSections"
        };

        private const string ToolDecisionSystemPrompt =
@"You are a read-only tool router for an Excel CSI toolbox.
Decide whether the user's message requires a local read-only tool.

Available read-only tools:
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

Return JSON only:
{
  ""shouldCallTool"": false,
  ""toolName"": """",
  ""argumentsJson"": ""{}"",
  ""reason"": """"
}

Strict rules:
- Only choose tools from the available read-only tool list.
- Never choose or invent write tools.
- If the user asks to modify the model, assign sections, assign loads, add/delete objects,
  run analysis, unlock model, or save model, set shouldCallTool = false and explain in reason
  that write operations are not allowed in this demo.
- Do not invent live model data.";

        private const string NormalChatSystemPrompt =
@"You are an AI assistant embedded inside an Excel CSI toolbox.
You can help the user understand Excel, ETABS, SAP2000, C# code, and structural engineering workflows.
In this demo version, you may retrieve real model information only through approved read-only local tools.
You cannot modify the ETABS/SAP2000/Excel model.
If the user asks you to modify the model, politely say this demo is read-only and cannot perform write actions.
Keep responses concise and practical.";

        private const string ToolResultSummarySystemPrompt =
@"You are an AI assistant inside an Excel CSI toolbox.
The user asked a question.
A read-only local tool was called and returned structured JSON.
Summarize the result clearly for the user.
Do not claim you performed any model modification.
If the result indicates failure, explain the failure simply and suggest checking whether
ETABS/SAP2000 is open and a model is loaded.";

        private readonly OllamaChatService _ollamaChatService;
        private readonly LocalMcpClient _mcpClient;

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

            if (!ApprovedReadOnlyTools.Contains(decision.ToolName))
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
                    AssistantText = "This demo is read-only and cannot modify the model. " + refusalText,
                    ToolWasCalled = false,
                    RoutingReason = $"Rejected non-approved tool: {decision.ToolName}"
                };
            }

            ToolCallResponse toolResponse = await _mcpClient.CallToolAsync(
                decision.ToolName,
                decision.ArgumentsJson,
                cancellationToken);

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

            if (ContainsAny(normalized, "selected frame", "selected member", "frame selected", "selected beam", "selected column"))
            {
                return CreateToolDecision("CSI.GetSelectedFrames", "Heuristic route: selected frames query.");
            }

            if (ContainsAny(normalized, "selected object", "current selection", "objects selected"))
            {
                return CreateToolDecision("CSI.GetSelectedObjects", "Heuristic route: current selection query.");
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
