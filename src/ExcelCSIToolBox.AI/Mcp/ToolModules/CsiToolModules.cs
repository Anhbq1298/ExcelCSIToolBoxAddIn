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

namespace ExcelCSIToolBox.AI.Mcp.ToolModules
{
    public static class ModelToolModule
    {
        public static void Register(IMcpToolRegistry registry, ToolServices services)
        {
            registry.Register(new CsiGetModelInfoTool(services.ReadOnlyConnectionService));
            registry.Register(new CsiGetModelStatisticsTool(services.EtabsService, services.Sap2000Service));
            registry.Register(new CsiRefreshViewTool(services.EtabsService, services.Sap2000Service));
            registry.Register(new CsiGetPresentUnitsTool(services.ReadOnlyConnectionService));
            registry.Register(new CsiGetSelectedObjectsTool(services.ReadOnlySelectionService));
            registry.Register(new SelectionClearTool(services.CommandService));
            registry.Register(new FileSaveModelTool(services.CommandService));
        }
    }

    public static class PointToolModule
    {
        public static void Register(IMcpToolRegistry registry, ToolServices services)
        {
            registry.Register(new PointsGetAllNamesTool(services.EtabsService, services.Sap2000Service));
            registry.Register(new PointsGetCountTool(services.EtabsService, services.Sap2000Service));
            registry.Register(new PointsGetByNameTool(services.EtabsService, services.Sap2000Service));
            registry.Register(new PointsGetCoordinatesTool(services.EtabsService, services.Sap2000Service));
            registry.Register(new PointsGetSelectedTool(services.EtabsService, services.Sap2000Service));
            registry.Register(new PointsGetRestraintTool(services.EtabsService, services.Sap2000Service));
            registry.Register(new PointsGetLoadForcesTool(services.EtabsService, services.Sap2000Service));
            registry.Register(new PointsGetSelectedByNameTool(services.EtabsService, services.Sap2000Service));
            registry.Register(new PointsGetGuidTool(services.EtabsService, services.Sap2000Service));
            registry.Register(new PointsGetGroupAssignmentsTool(services.EtabsService, services.Sap2000Service));
            registry.Register(new PointsGetConnectivityTool(services.EtabsService, services.Sap2000Service));
            registry.Register(new PointsGetSpringTool(services.EtabsService, services.Sap2000Service));
            registry.Register(new PointsGetMassTool(services.EtabsService, services.Sap2000Service));
            registry.Register(new PointsGetLocalAxesTool(services.EtabsService, services.Sap2000Service));
            registry.Register(new PointsGetDiaphragmTool(services.EtabsService, services.Sap2000Service));
            registry.Register(new PointsSetSelectedTool(services.EtabsService, services.Sap2000Service));
            registry.Register(new PointsAddByCoordinatesTool(services.CommandService));
            registry.Register(new PointsAddCartesianTool(services.CommandService));
            registry.Register(new PointsSetRestraintTool(services.EtabsService, services.Sap2000Service, services.WriteGuard, services.OperationLogger));
            registry.Register(new PointsSetLoadForceTool(services.EtabsService, services.Sap2000Service, services.WriteGuard, services.OperationLogger));
        }
    }

    public static class FrameToolModule
    {
        public static void Register(IMcpToolRegistry registry, ToolServices services)
        {
            registry.Register(new CsiGetSelectedFramesTool(services.ToolCatalogService));
            registry.Register(new CsiGetSelectedFrameSectionsTool(services.ReadOnlyFrameService));
            registry.Register(new FramesGetAllNamesTool(services.EtabsService, services.Sap2000Service));
            registry.Register(new FramesGetCountTool(services.EtabsService, services.Sap2000Service));
            registry.Register(new FramesGetByNameTool(services.EtabsService, services.Sap2000Service));
            registry.Register(new FramesGetPointsTool(services.EtabsService, services.Sap2000Service));
            registry.Register(new FramesGetSectionTool(services.EtabsService, services.Sap2000Service));
            registry.Register(new FramesGetDistributedLoadsTool(services.EtabsService, services.Sap2000Service));
            registry.Register(new FramesGetPointLoadsTool(services.EtabsService, services.Sap2000Service));
            registry.Register(new FramesGetSelectedTool(services.EtabsService, services.Sap2000Service));
            registry.Register(new FramesSetSelectedTool(services.EtabsService, services.Sap2000Service));
            registry.Register(new FramesGetSectionsTool(services.EtabsService, services.Sap2000Service));
            registry.Register(new FramesGetSectionDetailTool(services.EtabsService, services.Sap2000Service));
            registry.Register(new FramesAddByCoordinatesTool(services.CommandService));
            registry.Register(new FramesAddByPointsTool(services.CommandService));
            registry.Register(new FramesAddObjectTool(services.EtabsService, services.Sap2000Service));
            registry.Register(new FramesAddObjectsTool(services.EtabsService, services.Sap2000Service));
            registry.Register(new FramesAssignSectionTool(services.CommandService));
            registry.Register(new LoadsFrameAssignDistributedTool(services.CommandService));
            registry.Register(new LoadsFrameAssignPointLoadTool(services.CommandService));
            registry.Register(new FramesAssignDistributedLoadTool(services.CommandService));
            registry.Register(new FramesAssignPointLoadTool(services.CommandService));
            registry.Register(new FramesDeleteTool(services.CommandService));
        }
    }

    public static class ShellToolModule
    {
        public static void Register(IMcpToolRegistry registry, ToolServices services)
        {
            registry.Register(new CsiGetShellNamesTool(services.EtabsService, services.Sap2000Service));
            registry.Register(new ShellsGetCountTool(services.EtabsService, services.Sap2000Service));
            registry.Register(new ShellsGetByNameTool(services.EtabsService, services.Sap2000Service));
            registry.Register(new ShellsGetPointsTool(services.EtabsService, services.Sap2000Service));
            registry.Register(new ShellsGetPropertyTool(services.EtabsService, services.Sap2000Service));
            registry.Register(new ShellsGetSelectedTool(services.EtabsService, services.Sap2000Service));
            registry.Register(new ShellsGetUniformLoadsTool(services.EtabsService, services.Sap2000Service));
            registry.Register(new ShellsAddByPointsTool(services.EtabsService, services.Sap2000Service));
            registry.Register(new ShellsAddByCoordinatesTool(services.EtabsService, services.Sap2000Service));
            registry.Register(new ShellsAssignUniformLoadTool(services.EtabsService, services.Sap2000Service));
            registry.Register(new ShellsDeleteTool(services.EtabsService, services.Sap2000Service));
        }
    }

    public static class LoadToolModule
    {
        public static void Register(IMcpToolRegistry registry, ToolServices services)
        {
            registry.Register(new LoadCombinationsGetAllTool(services.EtabsService, services.Sap2000Service));
            registry.Register(new LoadCombinationsGetDetailsTool(services.EtabsService, services.Sap2000Service));
            registry.Register(new LoadCombinationsDeleteTool(services.EtabsService, services.Sap2000Service, services.WriteGuard, services.OperationLogger));
            registry.Register(new LoadPatternsGetAllTool(services.EtabsService, services.Sap2000Service));
            registry.Register(new LoadPatternsDeleteTool(services.EtabsService, services.Sap2000Service, services.WriteGuard, services.OperationLogger));
        }
    }

    public static class AnalysisToolModule
    {
        public static void Register(IMcpToolRegistry registry, ToolServices services)
        {
            registry.Register(new AnalyzeSelectedFramesTool(services.EtabsService, services.Sap2000Service));
            registry.Register(new AnalysisRunTool(services.CommandService));
        }
    }

    public static class TrussToolModule
    {
        public static void Register(IMcpToolRegistry registry, ToolServices services)
        {
            registry.Register(new TrussGenerateHoweTool(services.EtabsService, services.Sap2000Service, services.TrussGenerationService));
            registry.Register(new TrussGeneratePrattTool(services.EtabsService, services.Sap2000Service, services.TrussGenerationService));
        }
    }

    public static class WorkflowToolModule
    {
        public static void Register(IMcpToolRegistry registry, ToolServices services)
        {
            registry.Register(new ExecuteCsiRequestTool(services.EtabsService, services.Sap2000Service, services.WorkflowExecutionService));
            registry.Register(new RandomGenerateObjectsTool(services.EtabsService, services.Sap2000Service, services.RandomObjectGenerationService));
        }
    }
}
