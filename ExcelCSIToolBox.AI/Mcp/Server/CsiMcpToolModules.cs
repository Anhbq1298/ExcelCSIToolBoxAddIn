using ExcelCSIToolBox.AI.Mcp.Tools;
using ExcelCSIToolBox.AI.Mcp.Tools.CSI.Frames;
using ExcelCSIToolBox.AI.Mcp.Tools.CSI.Loads.LoadCombinations;
using ExcelCSIToolBox.AI.Mcp.Tools.CSI.Loads.LoadPatterns;
using ExcelCSIToolBox.AI.Mcp.Tools.CSI.Model;
using ExcelCSIToolBox.AI.Mcp.Tools.CSI.Points;
using ExcelCSIToolBox.AI.Mcp.Tools.CSI.Random;
using ExcelCSIToolBox.AI.Mcp.Tools.CSI.Shells;
using ExcelCSIToolBox.AI.Mcp.Tools.CSI.Truss;
using ExcelCSIToolBox.AI.Mcp.Tools.CSI.Workflow;

namespace ExcelCSIToolBox.AI.Mcp.Server
{
    public interface ICsiMcpToolModule
    {
        void Register(IMcpToolRegistry registry, CsiMcpToolContext context);
    }

    public sealed class CsiModelToolModule : ICsiMcpToolModule
    {
        public void Register(IMcpToolRegistry registry, CsiMcpToolContext context)
        {
            registry.Register(new CsiGetModelInfoTool(context.ReadOnlyConnectionService));
            registry.Register(new CsiGetModelStatisticsTool(context.EtabsService, context.Sap2000Service));
            registry.Register(new CsiRefreshViewTool(context.EtabsService, context.Sap2000Service));
            registry.Register(new CsiGetPresentUnitsTool(context.ReadOnlyConnectionService));
            registry.Register(new CsiGetSelectedObjectsTool(context.ReadOnlySelectionService));
            registry.Register(new SelectionClearTool(context.CommandService));
            registry.Register(new FileSaveModelTool(context.CommandService));
        }
    }

    public sealed class PointToolModule : ICsiMcpToolModule
    {
        public void Register(IMcpToolRegistry registry, CsiMcpToolContext context)
        {
            registry.Register(new PointsGetAllNamesTool(context.EtabsService, context.Sap2000Service));
            registry.Register(new PointsGetCountTool(context.EtabsService, context.Sap2000Service));
            registry.Register(new PointsGetByNameTool(context.EtabsService, context.Sap2000Service));
            registry.Register(new PointsGetCoordinatesTool(context.EtabsService, context.Sap2000Service));
            registry.Register(new PointsGetSelectedTool(context.EtabsService, context.Sap2000Service));
            registry.Register(new PointsGetRestraintTool(context.EtabsService, context.Sap2000Service));
            registry.Register(new PointsGetLoadForcesTool(context.EtabsService, context.Sap2000Service));
            registry.Register(new PointsGetSelectedByNameTool(context.EtabsService, context.Sap2000Service));
            registry.Register(new PointsGetGuidTool(context.EtabsService, context.Sap2000Service));
            registry.Register(new PointsGetGroupAssignmentsTool(context.EtabsService, context.Sap2000Service));
            registry.Register(new PointsGetConnectivityTool(context.EtabsService, context.Sap2000Service));
            registry.Register(new PointsGetSpringTool(context.EtabsService, context.Sap2000Service));
            registry.Register(new PointsGetMassTool(context.EtabsService, context.Sap2000Service));
            registry.Register(new PointsGetLocalAxesTool(context.EtabsService, context.Sap2000Service));
            registry.Register(new PointsGetDiaphragmTool(context.EtabsService, context.Sap2000Service));
            registry.Register(new PointsSetSelectedTool(context.EtabsService, context.Sap2000Service));
            registry.Register(new PointsAddByCoordinatesTool(context.CommandService));
            registry.Register(new PointsAddCartesianTool(context.CommandService));
            registry.Register(new PointsSetRestraintTool(context.EtabsService, context.Sap2000Service, context.WriteGuard, context.OperationLogger));
            registry.Register(new PointsSetLoadForceTool(context.EtabsService, context.Sap2000Service, context.WriteGuard, context.OperationLogger));
        }
    }

    public sealed class FrameToolModule : ICsiMcpToolModule
    {
        public void Register(IMcpToolRegistry registry, CsiMcpToolContext context)
        {
            registry.Register(new CsiGetSelectedFramesTool(context.ReadOnlySelectionService));
            registry.Register(new CsiGetSelectedFrameSectionsTool(context.ReadOnlyFrameService));
            registry.Register(new FramesGetAllNamesTool(context.EtabsService, context.Sap2000Service));
            registry.Register(new FramesGetCountTool(context.EtabsService, context.Sap2000Service));
            registry.Register(new FramesGetByNameTool(context.EtabsService, context.Sap2000Service));
            registry.Register(new FramesGetPointsTool(context.EtabsService, context.Sap2000Service));
            registry.Register(new FramesGetSectionTool(context.EtabsService, context.Sap2000Service));
            registry.Register(new FramesGetDistributedLoadsTool(context.EtabsService, context.Sap2000Service));
            registry.Register(new FramesGetPointLoadsTool(context.EtabsService, context.Sap2000Service));
            registry.Register(new FramesGetSelectedTool(context.EtabsService, context.Sap2000Service));
            registry.Register(new FramesSetSelectedTool(context.EtabsService, context.Sap2000Service));
            registry.Register(new FramesGetSectionsTool(context.EtabsService, context.Sap2000Service));
            registry.Register(new FramesGetSectionDetailTool(context.EtabsService, context.Sap2000Service));
            registry.Register(new FramesAddByCoordinatesTool(context.CommandService));
            registry.Register(new FramesAddByPointsTool(context.CommandService));
            registry.Register(new FramesAddObjectTool(context.EtabsService, context.Sap2000Service));
            registry.Register(new FramesAddObjectsTool(context.EtabsService, context.Sap2000Service));
            registry.Register(new FramesAssignSectionTool(context.CommandService));
            registry.Register(new LoadsFrameAssignDistributedTool(context.CommandService));
            registry.Register(new LoadsFrameAssignPointLoadTool(context.CommandService));
            registry.Register(new FramesAssignDistributedLoadTool(context.CommandService));
            registry.Register(new FramesAssignPointLoadTool(context.CommandService));
            registry.Register(new FramesDeleteTool(context.CommandService));
        }
    }

    public sealed class ShellToolModule : ICsiMcpToolModule
    {
        public void Register(IMcpToolRegistry registry, CsiMcpToolContext context)
        {
            registry.Register(new CsiGetShellNamesTool(context.EtabsService, context.Sap2000Service));
            registry.Register(new ShellsGetCountTool(context.EtabsService, context.Sap2000Service));
            registry.Register(new ShellsGetByNameTool(context.EtabsService, context.Sap2000Service));
            registry.Register(new ShellsGetPointsTool(context.EtabsService, context.Sap2000Service));
            registry.Register(new ShellsGetPropertyTool(context.EtabsService, context.Sap2000Service));
            registry.Register(new ShellsGetSelectedTool(context.EtabsService, context.Sap2000Service));
            registry.Register(new ShellsGetUniformLoadsTool(context.EtabsService, context.Sap2000Service));
            registry.Register(new ShellsAddByPointsTool(context.EtabsService, context.Sap2000Service));
            registry.Register(new ShellsAddByCoordinatesTool(context.EtabsService, context.Sap2000Service));
            registry.Register(new ShellsAssignUniformLoadTool(context.EtabsService, context.Sap2000Service));
            registry.Register(new ShellsDeleteTool(context.EtabsService, context.Sap2000Service));
        }
    }

    public sealed class LoadToolModule : ICsiMcpToolModule
    {
        public void Register(IMcpToolRegistry registry, CsiMcpToolContext context)
        {
            registry.Register(new LoadCombinationsGetAllTool(context.EtabsService, context.Sap2000Service));
            registry.Register(new LoadCombinationsGetDetailsTool(context.EtabsService, context.Sap2000Service));
            registry.Register(new LoadCombinationsDeleteTool(context.EtabsService, context.Sap2000Service, context.WriteGuard, context.OperationLogger));
            registry.Register(new LoadPatternsGetAllTool(context.EtabsService, context.Sap2000Service));
            registry.Register(new LoadPatternsDeleteTool(context.EtabsService, context.Sap2000Service, context.WriteGuard, context.OperationLogger));
        }
    }

    public sealed class TrussToolModule : ICsiMcpToolModule
    {
        public void Register(IMcpToolRegistry registry, CsiMcpToolContext context)
        {
            registry.Register(new TrussGenerateHoweTool(context.EtabsService, context.Sap2000Service, context.TrussGenerationService));
            registry.Register(new TrussGeneratePrattTool(context.EtabsService, context.Sap2000Service, context.TrussGenerationService));
        }
    }

    public sealed class WorkflowToolModule : ICsiMcpToolModule
    {
        public void Register(IMcpToolRegistry registry, CsiMcpToolContext context)
        {
            registry.Register(new ExecuteCsiRequestTool(context.EtabsService, context.Sap2000Service, context.WorkflowExecutionService));
            registry.Register(new RandomGenerateObjectsTool(context.EtabsService, context.Sap2000Service, context.RandomObjectGenerationService));
        }
    }

    public sealed class AnalysisToolModule : ICsiMcpToolModule
    {
        public void Register(IMcpToolRegistry registry, CsiMcpToolContext context)
        {
            registry.Register(new AnalyzeSelectedFramesTool(context.EtabsService, context.Sap2000Service));
            registry.Register(new AnalysisRunTool(context.CommandService));
        }
    }
}
