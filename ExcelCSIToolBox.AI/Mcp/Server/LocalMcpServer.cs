using System;
using System.Threading;
using System.Threading.Tasks;
using ExcelCSIToolBox.AI.Mcp.Contracts;
using ExcelCSIToolBox.AI.Mcp.Tools;
using ExcelCSIToolBox.Core.Abstractions.CSI;
using ExcelCSIToolBox.AI.Mcp.Tools.CSI;

namespace ExcelCSIToolBox.AI.Mcp.Server
{
    /// <summary>
    /// Local MCP server that owns the tool registry and exposes a safe CallToolAsync method.
    ///
    /// Safety rules enforced here:
    /// - Only read-only tools may be registered (enforced by McpToolRegistry).
    /// - If a requested tool is not found, a failed response is returned.
    /// - If a tool's IsReadOnly flag is somehow false, execution is refused.
    /// </summary>
    public class LocalMcpServer
    {
        private readonly IMcpToolRegistry _registry;

        /// <summary>
        /// Create the server and register all approved read-only CSI tools.
        /// </summary>
        public LocalMcpServer(
            ICsiReadOnlyConnectionService connectionService,
            ICsiReadOnlySelectionService  selectionService,
            ICsiReadOnlyFrameService      frameService)
        {
            _registry = new McpToolRegistry();

            // Register all approved read-only tools.
            // The registry itself will throw if any tool reports IsReadOnly = false.
            _registry.Register(new CsiGetModelInfoTool(connectionService));
            _registry.Register(new CsiGetPresentUnitsTool(connectionService));
            _registry.Register(new CsiGetSelectedObjectsTool(selectionService));
            _registry.Register(new CsiGetSelectedFramesTool(selectionService));
            _registry.Register(new CsiGetSelectedFrameSectionsTool(frameService));
        }

        /// <summary>
        /// Execute a tool by name and return a structured response.
        /// Returns a failure response if the tool is not found or not read-only.
        /// </summary>
        public async Task<ToolCallResponse> CallToolAsync(
            ToolCallRequest    request,
            CancellationToken  cancellationToken)
        {
            if (request == null || string.IsNullOrWhiteSpace(request.ToolName))
            {
                return new ToolCallResponse
                {
                    ToolName   = string.Empty,
                    Success    = false,
                    Message    = "Tool call request is missing a tool name.",
                    ResultJson = null
                };
            }

            IMcpTool tool = _registry.GetTool(request.ToolName);

            if (tool == null)
            {
                return new ToolCallResponse
                {
                    ToolName   = request.ToolName,
                    Success    = false,
                    Message    = $"Tool '{request.ToolName}' is not registered.",
                    ResultJson = null
                };
            }

            // Double-check: refuse any non-read-only tool even if somehow registered.
            if (!tool.IsReadOnly)
            {
                return new ToolCallResponse
                {
                    ToolName   = request.ToolName,
                    Success    = false,
                    Message    = $"Tool '{request.ToolName}' is not read-only and cannot be executed in this demo.",
                    ResultJson = null
                };
            }

            try
            {
                return await tool.ExecuteAsync(request.ArgumentsJson ?? "{}", cancellationToken);
            }
            catch (Exception ex)
            {
                return new ToolCallResponse
                {
                    ToolName   = request.ToolName,
                    Success    = false,
                    Message    = $"Tool '{request.ToolName}' threw an unexpected exception: {ex.Message}",
                    ResultJson = null
                };
            }
        }

        /// <summary>Returns the registry so the orchestrator can inspect available tools.</summary>
        public IMcpToolRegistry Registry => _registry;
    }
}
