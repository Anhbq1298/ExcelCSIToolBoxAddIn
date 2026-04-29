using ExcelCSIToolBox.AI.Mcp.Contracts;
using ExcelCSIToolBox.AI.Mcp.Tools.CSI.Base;
using ExcelCSIToolBox.Core.Abstractions.CSI;
using ExcelCSIToolBox.Core.Common.Results;
using ExcelCSIToolBox.Core.Models.CSI;
using ExcelCSIToolBox.Data.CSISapModel.Workflow;
using ExcelCSIToolBox.Infrastructure.CSISapModel.Workflow;
using Newtonsoft.Json;

namespace ExcelCSIToolBox.AI.Mcp.Tools.CSI.Workflow
{
    public sealed class ExecuteCsiRequestTool : CsiActiveConnectionToolBase
    {
        private readonly CsiWorkflowExecutionService _workflowService = new CsiWorkflowExecutionService();

        public ExecuteCsiRequestTool(
            ICSISapModelConnectionService etabsService,
            ICSISapModelConnectionService sap2000Service)
            : base(etabsService, sap2000Service)
        {
        }

        public override string Name => "execute_csi_request";
        public override string Title => "Execute CSI Request";
        public override string Category => "Workflow";
        public override string SubCategory => "Orchestration";
        public override string Description => "Executes a raw multi-step CSI request by parsing, ordering, and running supported tasks.";
        public override bool IsReadOnly => false;
        public override CsiMethodRiskLevel RiskLevel => CsiMethodRiskLevel.Medium;
        public override bool RequiresConfirmation => false;
        public override bool SupportsDryRun => false;

        protected override ToolCallResponse Execute(ICSISapModelConnectionService service, string argumentsJson)
        {
            CsiWorkflowRequestDto request = ReadArgs<CsiWorkflowRequestDto>(argumentsJson);
            OperationResult<CsiWorkflowResultDto> result = _workflowService.Execute(service, request);
            if (!result.IsSuccess)
            {
                return Fail(result.Message);
            }

            return new ToolCallResponse
            {
                ToolName = Name,
                Success = result.Data.Failed == 0 && result.Data.Skipped == 0,
                Message = result.Data.Failed == 0 && result.Data.Skipped == 0 ? "Success" : "Failure",
                ResultJson = JsonConvert.SerializeObject(result.Data)
            };
        }
    }
}
