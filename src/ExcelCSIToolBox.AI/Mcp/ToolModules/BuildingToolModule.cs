using ExcelCSIToolBox.AI.Mcp.Tools;
using ExcelCSIToolBox.AI.Mcp.Tools.Building;

namespace ExcelCSIToolBox.AI.Mcp.ToolModules
{
    public static class BuildingToolModule
    {
        public static void Register(IMcpToolRegistry registry, ToolServices services)
        {
            registry.Register(new BuildingGenerateOptionsTool(services.BuildingOptionService, services.ConstraintValidationService, services.ResultEvaluationService, services.OptionRankingService));
            registry.Register(new BuildingPreviewOptionTool(services.BuildingOptionService, services.ConstraintValidationService, services.ResultEvaluationService, services.OptionRankingService));
            registry.Register(new BuildingBuildOptionTool(services.BuildingOptionService, services.ConstraintValidationService, services.ResultEvaluationService, services.OptionRankingService));
            registry.Register(new BuildingRunAnalysisTool(services.BuildingOptionService, services.ConstraintValidationService, services.ResultEvaluationService, services.OptionRankingService));
            registry.Register(new BuildingEvaluateOptionTool(services.BuildingOptionService, services.ConstraintValidationService, services.ResultEvaluationService, services.OptionRankingService));
            registry.Register(new BuildingRankOptionsTool(services.BuildingOptionService, services.ConstraintValidationService, services.ResultEvaluationService, services.OptionRankingService));
        }
    }
}
