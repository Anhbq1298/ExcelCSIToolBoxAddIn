using System;
using System.Threading;
using System.Threading.Tasks;
using ExcelCSIToolBox.AI.Mcp.Contracts;
using ExcelCSIToolBox.AI.Mcp.Server;

namespace ExcelCSIToolBox.AI.Mcp.Client
{
    /// <summary>
    /// Thin client that delegates tool calls to the LocalMcpServer.
    /// The AI agent orchestrator uses this to call tools without knowing
    /// implementation details of the server or the tools themselves.
    /// </summary>
    public class LocalMcpClient
    {
        private readonly LocalMcpServer _server;

        public LocalMcpClient(LocalMcpServer server)
        {
            _server = server ?? throw new ArgumentNullException(nameof(server));
        }

        /// <summary>
        /// Forward a tool call to the server and return the response.
        /// </summary>
        public Task<ToolCallResponse> CallToolAsync(
            string            toolName,
            string            argumentsJson,
            CancellationToken cancellationToken)
        {
            ToolCallRequest request = new ToolCallRequest
            {
                ToolName      = toolName,
                ArgumentsJson = argumentsJson ?? "{}"
            };

            return _server.CallToolAsync(request, cancellationToken);
        }
    }
}
