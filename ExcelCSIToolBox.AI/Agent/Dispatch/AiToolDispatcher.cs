using System;
using System.Diagnostics;
using System.Threading;
using System.Threading.Tasks;
using ExcelCSIToolBox.AI.Mcp.Client;
using ExcelCSIToolBox.AI.Mcp.Contracts;
using Newtonsoft.Json.Linq;

namespace ExcelCSIToolBox.AI.Agent
{
    public sealed class AiToolDispatcher : IAiToolDispatcher
    {
        private readonly LocalMcpClient _mcpClient;

        public AiToolDispatcher(LocalMcpClient mcpClient)
        {
            _mcpClient = mcpClient ?? throw new ArgumentNullException(nameof(mcpClient));
        }

        public Task<ToolCallResponse> DispatchAsync(AiAgentToolDecision decision, CancellationToken cancellationToken)
        {
            if (decision == null || string.IsNullOrWhiteSpace(decision.ToolName))
            {
                string expectedToolName = decision == null || string.IsNullOrWhiteSpace(decision.MissingSchemaMessage)
                    ? "unknown"
                    : decision.MissingSchemaMessage.Replace("No MCP tool matched. Missing tool schema:", string.Empty).Trim().TrimEnd('.');

                string missingSchemaMessage = decision == null || string.IsNullOrWhiteSpace(decision.MissingSchemaMessage)
                    ? "No MCP tool matched. Missing tool schema: " + expectedToolName + "."
                    : decision.MissingSchemaMessage;

                var diagnostic = new JObject
                {
                    ["status"] = "NoToolMatched",
                    ["candidateDomain"] = decision == null ? null : decision.CandidateDomain,
                    ["candidateAction"] = decision == null ? null : decision.CandidateAction,
                    ["targetObject"] = decision == null ? null : decision.TargetObject,
                    ["missingSchemaMessage"] = missingSchemaMessage
                };

                return Task.FromResult(new ToolCallResponse
                {
                    ToolName = string.Empty,
                    Success = false,
                    Message = missingSchemaMessage,
                    ResultJson = diagnostic.ToString(Newtonsoft.Json.Formatting.None)
                });
            }

            Trace.WriteLine("AI MCP dispatch tool: " + decision.ToolName);
            Trace.WriteLine("AI MCP dispatch arguments: " + (decision.ArgumentsJson ?? "{}"));

            return _mcpClient.CallToolAsync(decision.ToolName, decision.ArgumentsJson, cancellationToken);
        }
    }
}
