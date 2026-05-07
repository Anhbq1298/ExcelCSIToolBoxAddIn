using System.Text.RegularExpressions;
using ExcelCSIToolBox.AI.Mcp.Contracts;
using ExcelCSIToolBox.AI.Mcp.Tools.CSI.Base;
using ExcelCSIToolBox.Core.Abstractions.CSI;
using ExcelCSIToolBox.Core.Common.Results;
using ExcelCSIToolBox.Core.Models.CSI;
using ExcelCSIToolBox.Data.CSISapModel.Workflow;
using Newtonsoft.Json;

namespace ExcelCSIToolBox.AI.Mcp.Tools.CSI.Workflow
{
    [MutationTool]
    public sealed class ExecuteCsiRequestTool : CsiActiveConnectionToolBase
    {
        private readonly ICsiWorkflowExecutionService _workflowService;

        public ExecuteCsiRequestTool(
            ICSISapModelConnectionService etabsService,
            ICSISapModelConnectionService sap2000Service,
            ICsiWorkflowExecutionService workflowService)
            : base(etabsService, sap2000Service)
        {
            _workflowService = workflowService ?? throw new System.ArgumentNullException(nameof(workflowService));
        }

        public override string Name => "execute_csi_request";
        public override string Title => "Execute CSI Request";
        public override string Category => "Workflow";
        public override string SubCategory => "Orchestration";
        public override string Description => "Executes a raw multi-step CSI request by parsing, ordering, and running supported tasks.";
        public override bool IsReadOnly => false;
        public override CsiMethodRiskLevel RiskLevel => CsiMethodRiskLevel.Medium;
        public override bool RequiresConfirmation => true;
        public override bool SupportsDryRun => true;

        protected override ToolCallResponse Execute(ICSISapModelConnectionService service, string argumentsJson)
        {
            CsiWorkflowRequestDto request = ReadArgs<CsiWorkflowRequestDto>(argumentsJson);
            if (LikelyWrites(request) && (request.DryRun || !request.Confirmed))
            {
                return Ok("Preview: workflow execution requires confirmation. Confirm to proceed?", new
                {
                    OperationName = Name,
                    Summary = "Execute CSI workflow: " + (request.UserInput ?? "planned task list"),
                    RequiresConfirmation = true,
                    SupportsDryRun = true
                });
            }

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

        private static bool LikelyWrites(CsiWorkflowRequestDto request)
        {
            if (request == null)
            {
                return false;
            }

            if (request.PlannedTasks != null)
            {
                for (int i = 0; i < request.PlannedTasks.Count; i++)
                {
                    CsiTaskDto task = request.PlannedTasks[i];
                    if (task != null && Regex.IsMatch(task.Operation ?? string.Empty, "Add|Assign|Delete|Set|Select", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant))
                    {
                        return true;
                    }
                }
            }

            return Regex.IsMatch(
                request.UserInput ?? string.Empty,
                @"\b(add|create|draw|insert|assign|apply|set|delete|remove|select|run|save)\b",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
        }
    }
}
