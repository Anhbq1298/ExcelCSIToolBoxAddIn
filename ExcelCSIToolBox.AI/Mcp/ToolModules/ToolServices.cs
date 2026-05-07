using System;
using ExcelCSIToolBox.Application.GenerativeDesign;
using ExcelCSIToolBox.Application.ToolCatalog.Contracts;
using ExcelCSIToolBox.AI.Mcp.Safety;
using ExcelCSIToolBox.AI.Mcp.Server;
using ExcelCSIToolBox.Core.Abstractions.CSI;

namespace ExcelCSIToolBox.AI.Mcp.ToolModules
{
    public sealed class ToolServices
    {
        public ToolServices(CsiMcpToolContext context)
        {
            if (context == null)
            {
                throw new ArgumentNullException(nameof(context));
            }

            ReadOnlyConnectionService = context.ReadOnlyConnectionService;
            ReadOnlySelectionService = context.ReadOnlySelectionService;
            ReadOnlyFrameService = context.ReadOnlyFrameService;
            EtabsService = context.EtabsService;
            Sap2000Service = context.Sap2000Service;
            CommandService = context.CommandService;
            WriteGuard = context.WriteGuard;
            OperationLogger = context.OperationLogger;
            RandomObjectGenerationService = context.RandomObjectGenerationService;
            TrussGenerationService = context.TrussGenerationService;
            WorkflowExecutionService = context.WorkflowExecutionService;
            ToolCatalogService = context.ToolCatalogService;
            MutationGuard = context.MutationGuard;
            BuildingOptionService = context.BuildingOptionService;
            ConstraintValidationService = context.ConstraintValidationService;
            ResultEvaluationService = context.ResultEvaluationService;
            OptionRankingService = context.OptionRankingService;
        }

        public ICsiReadOnlyConnectionService ReadOnlyConnectionService { get; }
        public ICsiReadOnlySelectionService ReadOnlySelectionService { get; }
        public ICsiReadOnlyFrameService ReadOnlyFrameService { get; }
        public ICSISapModelConnectionService EtabsService { get; }
        public ICSISapModelConnectionService Sap2000Service { get; }
        public ICsiModelCommandService CommandService { get; }
        public IMcpWriteGuard WriteGuard { get; }
        public ICsiOperationLogger OperationLogger { get; }
        public ICsiRandomObjectGenerationService RandomObjectGenerationService { get; }
        public ICsiTrussGenerationService TrussGenerationService { get; }
        public ICsiWorkflowExecutionService WorkflowExecutionService { get; }
        public IToolCatalogService ToolCatalogService { get; }
        public IMutationGuard MutationGuard { get; }
        public BuildingOptionService BuildingOptionService { get; }
        public ConstraintValidationService ConstraintValidationService { get; }
        public ResultEvaluationService ResultEvaluationService { get; }
        public OptionRankingService OptionRankingService { get; }
    }
}
