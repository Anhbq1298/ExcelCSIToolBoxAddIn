using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using ExcelCSIToolBox.AI.Mcp.Contracts;
using ExcelCSIToolBox.AI.Mcp.Tools;
using ExcelCSIToolBox.Core.Abstractions.CSI;
using ExcelCSIToolBox.AI.Mcp.Tools.CSI;
using ExcelCSIToolBox.Infrastructure.CSISapModel;
using ExcelCSIToolBox.Infrastructure.CSISapModel.Adapters;
using ExcelCSIToolBox.Infrastructure.Etabs;
using ExcelCSIToolBox.Infrastructure.Sap2000;

namespace ExcelCSIToolBox.AI.Mcp.Server
{
    /// <summary>
    /// Local MCP server that owns the tool registry and exposes a safe CallToolAsync method.
    ///
    /// Safety rules enforced here:
    /// - If a requested tool is not found, a failed response is returned.
    /// - Write tools must be typed tools backed by ICsiModelCommandService and IMcpWriteGuard.
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
            : this(connectionService, selectionService, frameService, CreateDefaultCommandService())
        {
        }

        public LocalMcpServer(
            ICsiReadOnlyConnectionService connectionService,
            ICsiReadOnlySelectionService  selectionService,
            ICsiReadOnlyFrameService      frameService,
            ICsiModelCommandService       commandService)
        {
            _registry = new McpToolRegistry();

            // Register all approved read-only tools.
            // The registry itself will throw if any tool reports IsReadOnly = false.
            _registry.Register(new CsiGetModelInfoTool(connectionService));
            _registry.Register(new CsiGetPresentUnitsTool(connectionService));
            _registry.Register(new CsiGetSelectedObjectsTool(selectionService));
            _registry.Register(new CsiGetSelectedFramesTool(selectionService));
            _registry.Register(new CsiGetSelectedFrameSectionsTool(frameService));

            _registry.Register(new PointsAddByCoordinatesTool(commandService));
            _registry.Register(new FramesAddByCoordinatesTool(commandService));
            _registry.Register(new FramesAddByPointsTool(commandService));
            _registry.Register(new FramesAssignSectionTool(commandService));
            _registry.Register(new LoadsFrameAssignDistributedTool(commandService));
            _registry.Register(new LoadsFrameAssignPointLoadTool(commandService));
            _registry.Register(new SelectionClearTool(commandService));
            _registry.Register(new FramesDeleteTool(commandService));
            _registry.Register(new AnalysisRunTool(commandService));
            _registry.Register(new FileSaveModelTool(commandService));
        }

        /// <summary>
        /// Execute a tool by name and return a structured response.
        /// Returns a failure response if the tool is not found.
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

        public IReadOnlyList<McpToolDescriptor> ListTools()
        {
            var descriptors = new List<McpToolDescriptor>();
            IReadOnlyList<IMcpTool> tools = _registry.GetAllTools();

            for (int i = 0; i < tools.Count; i++)
            {
                IMcpTool tool = tools[i];
                IMcpToolMetadata metadata = tool as IMcpToolMetadata;

                descriptors.Add(new McpToolDescriptor
                {
                    Name = tool.Name,
                    Title = metadata == null ? tool.Name : metadata.Title,
                    Category = metadata == null ? "CSI" : metadata.Category,
                    SubCategory = metadata == null ? "Read" : metadata.SubCategory,
                    Description = tool.Description,
                    IsReadOnly = tool.IsReadOnly,
                    RiskLevel = metadata == null ? Core.Models.CSI.CsiMethodRiskLevel.None : metadata.RiskLevel,
                    RequiresConfirmation = metadata != null && metadata.RequiresConfirmation,
                    SupportsDryRun = metadata != null && metadata.SupportsDryRun
                });
            }

            return descriptors;
        }

        private static ICsiModelCommandService CreateDefaultCommandService()
        {
            return new CsiModelCommandService(
                new EtabsConnectionService(new EtabsModelAdapter()),
                new Sap2000ConnectionService(new Sap2000ModelAdapter()),
                new CsiWriteGuard(),
                new CsiOperationLogger());
        }
    }
}
