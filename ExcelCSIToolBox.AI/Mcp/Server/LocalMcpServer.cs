using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using ExcelCSIToolBox.AI.Mcp.Contracts;
using ExcelCSIToolBox.AI.Mcp.Tools;
using ExcelCSIToolBox.Core.Abstractions.CSI;
using ExcelCSIToolBox.AI.Mcp.Tools.CSI.Frames;
using ExcelCSIToolBox.AI.Mcp.Tools.CSI.Loads.LoadCombinations;
using ExcelCSIToolBox.AI.Mcp.Tools.CSI.Loads.LoadPatterns;
using ExcelCSIToolBox.AI.Mcp.Tools.CSI.Model;
using ExcelCSIToolBox.AI.Mcp.Tools.CSI.Points;
using ExcelCSIToolBox.AI.Mcp.Tools.CSI.Random;
using ExcelCSIToolBox.AI.Mcp.Tools.CSI.Shells;
using ExcelCSIToolBox.AI.Mcp.Tools.CSI.Truss;
using ExcelCSIToolBox.AI.Mcp.Tools.CSI.Workflow;
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
        private readonly SynchronizationContext _toolSynchronizationContext;

        /// <summary>
        /// Create the server and register all approved read-only CSI tools.
        /// </summary>
        public LocalMcpServer(
            ICsiReadOnlyConnectionService connectionService,
            ICsiReadOnlySelectionService  selectionService,
            ICsiReadOnlyFrameService      frameService)
            : this(
                connectionService,
                selectionService,
                frameService,
                new EtabsConnectionService(new EtabsModelAdapter()),
                new Sap2000ConnectionService(new Sap2000ModelAdapter()))
        {
        }

        private LocalMcpServer(
            ICsiReadOnlyConnectionService connectionService,
            ICsiReadOnlySelectionService  selectionService,
            ICsiReadOnlyFrameService      frameService,
            ICSISapModelConnectionService etabsService,
            ICSISapModelConnectionService sap2000Service)
            : this(
                connectionService,
                selectionService,
                frameService,
                new CsiModelCommandService(
                    etabsService,
                    sap2000Service,
                    new CsiWriteGuard(),
                    new CsiOperationLogger()),
                etabsService,
                sap2000Service)
        {
        }

        public LocalMcpServer(
            ICsiReadOnlyConnectionService connectionService,
            ICsiReadOnlySelectionService  selectionService,
            ICsiReadOnlyFrameService      frameService,
            ICsiModelCommandService       commandService)
            : this(
                connectionService,
                selectionService,
                frameService,
                commandService,
                new EtabsConnectionService(new EtabsModelAdapter()),
                new Sap2000ConnectionService(new Sap2000ModelAdapter()))
        {
        }

        private LocalMcpServer(
            ICsiReadOnlyConnectionService connectionService,
            ICsiReadOnlySelectionService  selectionService,
            ICsiReadOnlyFrameService      frameService,
            ICsiModelCommandService       commandService,
            ICSISapModelConnectionService etabsService,
            ICSISapModelConnectionService sap2000Service)
        {
            _toolSynchronizationContext = SynchronizationContext.Current;
            _registry = new McpToolRegistry();

            // Register all approved read-only tools.
            // The registry itself will throw if any tool reports IsReadOnly = false.
            _registry.Register(new CsiGetModelInfoTool(connectionService));
            _registry.Register(new CsiGetModelStatisticsTool(etabsService, sap2000Service));
            _registry.Register(new CsiRefreshViewTool(etabsService, sap2000Service));
            _registry.Register(new CsiGetPresentUnitsTool(connectionService));
            _registry.Register(new CsiGetSelectedObjectsTool(selectionService));
            _registry.Register(new CsiGetSelectedFramesTool(selectionService));
            _registry.Register(new CsiGetSelectedFrameSectionsTool(frameService));
            _registry.Register(new PointsGetAllNamesTool(etabsService, sap2000Service));
            _registry.Register(new PointsGetCountTool(etabsService, sap2000Service));
            _registry.Register(new PointsGetByNameTool(etabsService, sap2000Service));
            _registry.Register(new PointsGetCoordinatesTool(etabsService, sap2000Service));
            _registry.Register(new PointsGetSelectedTool(etabsService, sap2000Service));
            _registry.Register(new PointsGetRestraintTool(etabsService, sap2000Service));
            _registry.Register(new PointsGetLoadForcesTool(etabsService, sap2000Service));
            _registry.Register(new PointsGetSelectedByNameTool(etabsService, sap2000Service));
            _registry.Register(new PointsGetGuidTool(etabsService, sap2000Service));
            _registry.Register(new PointsGetGroupAssignmentsTool(etabsService, sap2000Service));
            _registry.Register(new PointsGetConnectivityTool(etabsService, sap2000Service));
            _registry.Register(new PointsGetSpringTool(etabsService, sap2000Service));
            _registry.Register(new PointsGetMassTool(etabsService, sap2000Service));
            _registry.Register(new PointsGetLocalAxesTool(etabsService, sap2000Service));
            _registry.Register(new PointsGetDiaphragmTool(etabsService, sap2000Service));
            _registry.Register(new PointsSetSelectedTool(etabsService, sap2000Service));
            _registry.Register(new FramesGetAllNamesTool(etabsService, sap2000Service));
            _registry.Register(new FramesGetCountTool(etabsService, sap2000Service));
            _registry.Register(new FramesGetByNameTool(etabsService, sap2000Service));
            _registry.Register(new FramesGetPointsTool(etabsService, sap2000Service));
            _registry.Register(new FramesGetSectionTool(etabsService, sap2000Service));
            _registry.Register(new FramesGetDistributedLoadsTool(etabsService, sap2000Service));
            _registry.Register(new FramesGetPointLoadsTool(etabsService, sap2000Service));
            _registry.Register(new FramesGetSelectedTool(etabsService, sap2000Service));
            _registry.Register(new FramesSetSelectedTool(etabsService, sap2000Service));
            _registry.Register(new FramesGetSectionsTool(etabsService, sap2000Service));
            _registry.Register(new FramesGetSectionDetailTool(etabsService, sap2000Service));
            _registry.Register(new AnalyzeSelectedFramesTool(etabsService, sap2000Service));
            _registry.Register(new CsiGetShellNamesTool(etabsService, sap2000Service));
            _registry.Register(new ShellsGetCountTool(etabsService, sap2000Service));
            _registry.Register(new ShellsGetByNameTool(etabsService, sap2000Service));
            _registry.Register(new ShellsGetPointsTool(etabsService, sap2000Service));
            _registry.Register(new ShellsGetPropertyTool(etabsService, sap2000Service));
            _registry.Register(new ShellsGetSelectedTool(etabsService, sap2000Service));
            _registry.Register(new ShellsGetUniformLoadsTool(etabsService, sap2000Service));
            _registry.Register(new ShellsAddByPointsTool(etabsService, sap2000Service));
            _registry.Register(new ShellsAddByCoordinatesTool(etabsService, sap2000Service));
            _registry.Register(new ShellsAssignUniformLoadTool(etabsService, sap2000Service));
            _registry.Register(new ShellsDeleteTool(etabsService, sap2000Service));
            _registry.Register(new LoadCombinationsGetAllTool(etabsService, sap2000Service));
            _registry.Register(new LoadCombinationsGetDetailsTool(etabsService, sap2000Service));
            _registry.Register(new LoadCombinationsDeleteTool(etabsService, sap2000Service));
            _registry.Register(new LoadPatternsGetAllTool(etabsService, sap2000Service));
            _registry.Register(new LoadPatternsDeleteTool(etabsService, sap2000Service));
            _registry.Register(new ExecuteCsiRequestTool(etabsService, sap2000Service));
            _registry.Register(new RandomGenerateObjectsTool(etabsService, sap2000Service));
            _registry.Register(new TrussGenerateHoweTool(etabsService, sap2000Service));
            _registry.Register(new TrussGeneratePrattTool(etabsService, sap2000Service));

            _registry.Register(new PointsAddByCoordinatesTool(commandService));
            _registry.Register(new PointsAddCartesianTool(commandService));
            _registry.Register(new PointsSetRestraintTool(etabsService, sap2000Service));
            _registry.Register(new PointsSetLoadForceTool(etabsService, sap2000Service));
            _registry.Register(new FramesAddByCoordinatesTool(commandService));
            _registry.Register(new FramesAddByPointsTool(commandService));
            _registry.Register(new FramesAddObjectTool(etabsService, sap2000Service));
            _registry.Register(new FramesAddObjectsTool(etabsService, sap2000Service));
            _registry.Register(new FramesAssignSectionTool(commandService));
            _registry.Register(new LoadsFrameAssignDistributedTool(commandService));
            _registry.Register(new LoadsFrameAssignPointLoadTool(commandService));
            _registry.Register(new FramesAssignDistributedLoadTool(commandService));
            _registry.Register(new FramesAssignPointLoadTool(commandService));
            _registry.Register(new SelectionClearTool(commandService));
            _registry.Register(new FramesDeleteTool(commandService));
            _registry.Register(new AnalysisRunTool(commandService));
            _registry.Register(new FileSaveModelTool(commandService));

            RegisterBackwardCompatibleToolAliases(_registry);
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

            return await ExecuteToolOnCapturedContextAsync(tool, request, cancellationToken);
        }

        /// <summary>Returns the registry so the orchestrator can inspect available tools.</summary>
        public IMcpToolRegistry Registry => _registry;

        private static void RegisterBackwardCompatibleToolAliases(IMcpToolRegistry registry)
        {
            registry.RegisterAlias("PointObj_AddCartesian", "points.add_by_coordinates");
            registry.RegisterAlias("PointObject_AddCartesian", "points.add_by_coordinates");
            registry.RegisterAlias("FrameObj_AddByPoint", "frames.add_object");
            registry.RegisterAlias("FrameObject_AddByPoint", "frames.add_object");
            registry.RegisterAlias("FrameObj_AddByCoordinate", "frames.add_object");
            registry.RegisterAlias("FrameObject_AddByCoordinate", "frames.add_object");
            registry.RegisterAlias("FrameObj_SetSection", "frames.assign_section");
            registry.RegisterAlias("FrameObject_AssignSection", "frames.assign_section");
            registry.RegisterAlias("SectionProperty_AssignToFrame", "frames.assign_section");
            registry.RegisterAlias("ShellObj_AddByPoint", "shells.add_by_points");
            registry.RegisterAlias("ShellObject_AddByPoint", "shells.add_by_points");
            registry.RegisterAlias("ShellObject_AddByCoordinate", "shells.add_by_coordinates");
            registry.RegisterAlias("Model_GetPresentUnits", "CSI.GetPresentUnits");
            registry.RegisterAlias("Model_GetFileName", "CSI.GetModelInfo");
            registry.RegisterAlias("Model_RefreshView", "csi.refresh_view");
            registry.RegisterAlias("Selection_Clear", "selection.clear");
            registry.RegisterAlias("Selection_GetSelectedObjects", "CSI.GetSelectedObjects");
            registry.RegisterAlias("LoadPattern_GetList", "loads.patterns.get_all");
            registry.RegisterAlias("LoadCombination_GetList", "loads.combinations.get_all");
            registry.RegisterAlias("Workflow_CreateTruss", "truss.generate_howe");
            registry.RegisterAlias("FrameObject_AssignDistributedLoad", "frames.assign_distributed_load");
        }

        private Task<ToolCallResponse> ExecuteToolOnCapturedContextAsync(
            IMcpTool tool,
            ToolCallRequest request,
            CancellationToken cancellationToken)
        {
            if (_toolSynchronizationContext == null ||
                SynchronizationContext.Current == _toolSynchronizationContext)
            {
                return ExecuteToolCoreAsync(tool, request, cancellationToken);
            }

            var completion = new TaskCompletionSource<ToolCallResponse>();
            _toolSynchronizationContext.Post(async state =>
            {
                try
                {
                    ToolCallResponse response = await ExecuteToolCoreAsync(tool, request, cancellationToken);
                    completion.TrySetResult(response);
                }
                catch (Exception ex)
                {
                    completion.TrySetResult(new ToolCallResponse
                    {
                        ToolName = request.ToolName,
                        Success = false,
                        Message = $"Tool '{request.ToolName}' threw an unexpected exception: {ex.Message}",
                        ResultJson = null
                    });
                }
            }, null);

            return completion.Task;
        }

        private static async Task<ToolCallResponse> ExecuteToolCoreAsync(
            IMcpTool tool,
            ToolCallRequest request,
            CancellationToken cancellationToken)
        {
            try
            {
                return await tool.ExecuteAsync(request.ArgumentsJson ?? "{}", cancellationToken);
            }
            catch (Exception ex)
            {
                return new ToolCallResponse
                {
                    ToolName = request.ToolName,
                    Success = false,
                    Message = $"Tool '{request.ToolName}' threw an unexpected exception: {ex.Message}",
                    ResultJson = null
                };
            }
        }

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

    }
}
