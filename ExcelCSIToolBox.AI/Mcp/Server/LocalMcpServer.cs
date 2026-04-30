using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using ExcelCSIToolBox.AI.Mcp.Contracts;
using ExcelCSIToolBox.AI.Mcp.ToolModules;
using ExcelCSIToolBox.AI.Mcp.Tools;
using ExcelCSIToolBox.Core.Models.CSI;

namespace ExcelCSIToolBox.AI.Mcp.Server
{
    /// <summary>
    /// Local MCP server that coordinates module registration and safe tool execution.
    /// Infrastructure dependencies are supplied by the add-in composition root.
    /// </summary>
    public class LocalMcpServer
    {
        private readonly IMcpToolRegistry _registry;
        private readonly SynchronizationContext _toolSynchronizationContext;

        public LocalMcpServer(CsiMcpToolContext context)
        {
            if (context == null)
            {
                throw new ArgumentNullException(nameof(context));
            }

            _toolSynchronizationContext = SynchronizationContext.Current;
            _registry = new McpToolRegistry();

            RegisterDefaultModules(_registry, new ToolServices(context));
            RegisterBackwardCompatibleToolAliases(_registry);
        }

        public IMcpToolRegistry Registry => _registry;

        public async Task<ToolCallResponse> CallToolAsync(
            ToolCallRequest request,
            CancellationToken cancellationToken)
        {
            if (request == null || string.IsNullOrWhiteSpace(request.ToolName))
            {
                return new ToolCallResponse
                {
                    ToolName = string.Empty,
                    Success = false,
                    Message = "Tool call request is missing a tool name.",
                    ResultJson = null
                };
            }

            IMcpTool tool = _registry.GetTool(request.ToolName);
            if (tool == null)
            {
                return new ToolCallResponse
                {
                    ToolName = request.ToolName,
                    Success = false,
                    Message = $"Tool '{request.ToolName}' is not registered.",
                    ResultJson = null
                };
            }

            return await ExecuteToolOnCapturedContextAsync(tool, request, cancellationToken);
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
                    RiskLevel = metadata == null ? CsiMethodRiskLevel.None : metadata.RiskLevel,
                    RequiresConfirmation = metadata != null && metadata.RequiresConfirmation,
                    SupportsDryRun = metadata != null && metadata.SupportsDryRun
                });
            }

            return descriptors;
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

            var completion = new TaskCompletionSource<ToolCallResponse>(TaskCreationOptions.RunContinuationsAsynchronously);
            CancellationTokenRegistration cancellationRegistration = cancellationToken.Register(
                () => completion.TrySetCanceled());
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
                finally
                {
                    cancellationRegistration.Dispose();
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

        private static void RegisterDefaultModules(IMcpToolRegistry registry, ToolServices services)
        {
            ModelToolModule.Register(registry, services);
            PointToolModule.Register(registry, services);
            FrameToolModule.Register(registry, services);
            ShellToolModule.Register(registry, services);
            LoadToolModule.Register(registry, services);
            AnalysisToolModule.Register(registry, services);
            TrussToolModule.Register(registry, services);
            WorkflowToolModule.Register(registry, services);
            BuildingToolModule.Register(registry, services);
        }

        private static void RegisterBackwardCompatibleToolAliases(IMcpToolRegistry registry)
        {
            RegisterPreferredAliases(registry);

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
            registry.RegisterAlias("Building_GenerateOptions", "building.generate_options");
            registry.RegisterAlias("Building_PreviewOption", "building.preview_option");
            registry.RegisterAlias("Building_BuildOption", "building.build_option");
            registry.RegisterAlias("Building_RunAnalysis", "building.run_analysis");
            registry.RegisterAlias("Building_EvaluateOption", "building.evaluate_option");
            registry.RegisterAlias("Building_RankOptions", "building.rank_options");
        }

        private static void RegisterPreferredAliases(IMcpToolRegistry registry)
        {
            registry.RegisterAlias("csi.model.get_info", "CSI.GetModelInfo");
            registry.RegisterAlias("csi.model.get_present_units", "CSI.GetPresentUnits");
            registry.RegisterAlias("csi.selection.get_selected_objects", "CSI.GetSelectedObjects");
            registry.RegisterAlias("csi.points.get_selected", "points.get_selected");
            registry.RegisterAlias("csi.points.add_by_coordinates", "points.add_by_coordinates");
            registry.RegisterAlias("csi.frames.add_by_coordinates", "frames.add_by_coordinates");
            registry.RegisterAlias("csi.frames.add_by_points", "frames.add_by_points");
            registry.RegisterAlias("csi.frames.assign_section", "frames.assign_section");
            registry.RegisterAlias("csi.frames.assign_distributed_load", "frames.assign_distributed_load");
            registry.RegisterAlias("csi.frames.assign_point_load", "frames.assign_point_load");
            registry.RegisterAlias("csi.shells.add_by_points", "shells.add_by_points");
            registry.RegisterAlias("csi.shells.add_by_coordinates", "shells.add_by_coordinates");
            registry.RegisterAlias("csi.loads.patterns.get_all", "loads.patterns.get_all");
            registry.RegisterAlias("csi.loads.combinations.get_all", "loads.combinations.get_all");
            registry.RegisterAlias("csi.truss.generate_howe", "truss.generate_howe");
            registry.RegisterAlias("csi.truss.generate_pratt", "truss.generate_pratt");
            registry.RegisterAlias("csi.frames.get_selected", "frames.get_selected");
            registry.RegisterAlias("csi.shells.get_selected", "shells.get_selected");
        }
    }
}
