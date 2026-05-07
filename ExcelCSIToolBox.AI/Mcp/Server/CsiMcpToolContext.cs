using System;
using ExcelCSIToolBox.Application.GenerativeDesign;
using ExcelCSIToolBox.Application.ToolCatalog.Contracts;
using ExcelCSIToolBox.AI.Mcp.Safety;
using ExcelCSIToolBox.Core.Abstractions.CSI;

namespace ExcelCSIToolBox.AI.Mcp.Server
{
    public sealed class CsiMcpToolContext
    {
        public CsiMcpToolContext(
            ICsiReadOnlyConnectionService readOnlyConnectionService,
            ICsiReadOnlySelectionService readOnlySelectionService,
            ICsiReadOnlyFrameService readOnlyFrameService,
            ICSISapModelConnectionService etabsService,
            ICSISapModelConnectionService sap2000Service,
            ICsiModelCommandService commandService,
            IMcpWriteGuard writeGuard,
            ICsiOperationLogger operationLogger,
            ICsiRandomObjectGenerationService randomObjectGenerationService,
            ICsiTrussGenerationService trussGenerationService,
            ICsiWorkflowExecutionService workflowExecutionService,
            IToolCatalogService toolCatalogService,
            IMutationGuard mutationGuard,
            BuildingOptionService buildingOptionService,
            ConstraintValidationService constraintValidationService,
            ResultEvaluationService resultEvaluationService,
            OptionRankingService optionRankingService)
        {
            ReadOnlyConnectionService = readOnlyConnectionService ?? throw new ArgumentNullException(nameof(readOnlyConnectionService));
            ReadOnlySelectionService = readOnlySelectionService ?? throw new ArgumentNullException(nameof(readOnlySelectionService));
            ReadOnlyFrameService = readOnlyFrameService ?? throw new ArgumentNullException(nameof(readOnlyFrameService));
            EtabsService = etabsService ?? throw new ArgumentNullException(nameof(etabsService));
            Sap2000Service = sap2000Service ?? throw new ArgumentNullException(nameof(sap2000Service));
            CommandService = commandService ?? throw new ArgumentNullException(nameof(commandService));
            WriteGuard = writeGuard ?? throw new ArgumentNullException(nameof(writeGuard));
            OperationLogger = operationLogger ?? throw new ArgumentNullException(nameof(operationLogger));
            RandomObjectGenerationService = randomObjectGenerationService ?? throw new ArgumentNullException(nameof(randomObjectGenerationService));
            TrussGenerationService = trussGenerationService ?? throw new ArgumentNullException(nameof(trussGenerationService));
            WorkflowExecutionService = workflowExecutionService ?? throw new ArgumentNullException(nameof(workflowExecutionService));
            ToolCatalogService = toolCatalogService ?? throw new ArgumentNullException(nameof(toolCatalogService));
            MutationGuard = mutationGuard ?? throw new ArgumentNullException(nameof(mutationGuard));
            BuildingOptionService = buildingOptionService ?? throw new ArgumentNullException(nameof(buildingOptionService));
            ConstraintValidationService = constraintValidationService ?? throw new ArgumentNullException(nameof(constraintValidationService));
            ResultEvaluationService = resultEvaluationService ?? throw new ArgumentNullException(nameof(resultEvaluationService));
            OptionRankingService = optionRankingService ?? throw new ArgumentNullException(nameof(optionRankingService));
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
