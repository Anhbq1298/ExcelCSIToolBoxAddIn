using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using ExcelCSIToolBox.AI.Mcp.Client;
using ExcelCSIToolBox.AI.Mcp.Contracts;
using ExcelCSIToolBox.AI.Ollama;

namespace ExcelCSIToolBox.AI.Agent
{
    /// <summary>
    /// Orchestrates the full AI agent conversation loop:
    ///
    ///   1. Ask Ollama whether a read-only tool should be called (tool-routing pass).
    ///   2. If yes: validate tool name, call LocalMcpClient, ask Ollama to summarise the result.
    ///   3. If no:  ask Ollama for a normal chat response.
    ///
    /// The AI cannot modify the model. If the LLM suggests a write tool, the orchestrator
    /// refuses and returns a read-only refusal message.
    /// </summary>
    public class AiAgentOrchestrator
    {
        // Approved read-only tool names. Any tool name not in this list is rejected.
        private static readonly HashSet<string> ApprovedReadOnlyTools = new HashSet<string>(
            StringComparer.OrdinalIgnoreCase)
        {
            "CSI.GetModelInfo",
            "CSI.GetPresentUnits",
            "CSI.GetSelectedObjects",
            "CSI.GetSelectedFrames",
            "CSI.GetSelectedFrameSections"
        };

        // ── System prompts ────────────────────────────────────────────────────────

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

        // ── Dependencies ──────────────────────────────────────────────────────────

        private readonly OllamaChatService _ollamaChatService;
        private readonly LocalMcpClient    _mcpClient;

        public AiAgentOrchestrator(OllamaChatService ollamaChatService, LocalMcpClient mcpClient)
        {
            _ollamaChatService = ollamaChatService
                ?? throw new ArgumentNullException(nameof(ollamaChatService));
            _mcpClient = mcpClient
                ?? throw new ArgumentNullException(nameof(mcpClient));
        }

        // ── Main entry point ──────────────────────────────────────────────────────

        /// <summary>
        /// Process one user message and return the agent's response.
        /// </summary>
        public async Task<AiAgentResponse> SendAsync(
            string            userMessage,
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

            // ── Pass 1: ask the LLM to decide if a tool call is needed ─────────────
            string decisionResponse = await _ollamaChatService.ChatAsync(
                new List<OllamaMessage>
                {
                    new OllamaMessage { role = "system",  content = ToolDecisionSystemPrompt },
                    new OllamaMessage { role = "user",    content = userMessage }
                },
                cancellationToken);

            AiAgentToolDecision decision = AiAgentToolDecisionParser.Parse(decisionResponse);

            // ── No tool call requested ────────────────────────────────────────────
            if (!decision.ShouldCallTool || string.IsNullOrWhiteSpace(decision.ToolName))
            {
                string chatText = await _ollamaChatService.ChatAsync(
                    new List<OllamaMessage>
                    {
                        new OllamaMessage { role = "system", content = NormalChatSystemPrompt },
                        new OllamaMessage { role = "user",   content = userMessage }
                    },
                    cancellationToken);

                return new AiAgentResponse
                {
                    AssistantText    = chatText,
                    ToolWasCalled    = false,
                    RoutingReason    = decision.Reason
                };
            }

            // ── Validate that the chosen tool is in the approved read-only list ────
            if (!ApprovedReadOnlyTools.Contains(decision.ToolName))
            {
                // The LLM proposed a tool that is not approved. Refuse and chat normally.
                string refusalText = await _ollamaChatService.ChatAsync(
                    new List<OllamaMessage>
                    {
                        new OllamaMessage { role = "system", content = NormalChatSystemPrompt },
                        new OllamaMessage { role = "user",   content = userMessage }
                    },
                    cancellationToken);

                return new AiAgentResponse
                {
                    AssistantText = "This demo is read-only and cannot modify the model. " + refusalText,
                    ToolWasCalled = false,
                    RoutingReason = $"Rejected non-approved tool: {decision.ToolName}"
                };
            }

            // ── Pass 2: call the approved read-only tool ──────────────────────────
            ToolCallResponse toolResponse = await _mcpClient.CallToolAsync(
                decision.ToolName,
                decision.ArgumentsJson,
                cancellationToken);

            // ── Pass 3: ask the LLM to summarize the tool result for the user ─────
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
                    new OllamaMessage { role = "user",   content = toolResultContext }
                },
                cancellationToken);

            return new AiAgentResponse
            {
                AssistantText    = summaryText,
                ToolWasCalled    = true,
                ToolName         = toolResponse.ToolName,
                ToolArgumentsJson = decision.ArgumentsJson,
                ToolResponse     = toolResponse,
                RoutingReason    = decision.Reason
            };
        }
    }
}
