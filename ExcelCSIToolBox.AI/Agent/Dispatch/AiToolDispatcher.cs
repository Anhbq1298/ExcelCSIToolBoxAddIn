using System;
using System.Threading;
using System.Threading.Tasks;
using ExcelCSIToolBox.AI.Mcp.Client;
using ExcelCSIToolBox.AI.Mcp.Contracts;

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
            if (decision == null)
            {
                throw new ArgumentNullException(nameof(decision));
            }

            return _mcpClient.CallToolAsync(decision.ToolName, decision.ArgumentsJson, cancellationToken);
        }
    }
}
