using System;
using System.Threading;
using System.Threading.Tasks;
using ExcelCSIToolBox.AI.Mcp.Contracts;
using ExcelCSIToolBox.AI.Mcp.Tools.CSI.Base;
using ExcelCSIToolBox.Core.Abstractions.CSI;
using ExcelCSIToolBox.Core.Models.CSI;

namespace ExcelCSIToolBox.AI.Mcp.Tools.CSI.Shells
{
    /// <summary>
    /// Read-only MCP tool: returns total shell/area object count from the active CSI model.
    /// </summary>
    public sealed class ShellsGetCountTool : CsiActiveConnectionToolBase
    {
        public ShellsGetCountTool(ICSISapModelConnectionService etabsService, ICSISapModelConnectionService sap2000Service) : base(etabsService, sap2000Service) { }
        public override string Name => "shells.get_count";
        public override string Title => "Get Shell Count";
        public override string Category => "Shells / Areas";
        public override string SubCategory => "Read";
        public override string Description => "Returns the total number of shell/area objects in the active CSI model.";
        public override bool IsReadOnly => true;
        public override CsiMethodRiskLevel RiskLevel => CsiMethodRiskLevel.None;
        public override bool RequiresConfirmation => false;
        public override bool SupportsDryRun => false;

        protected override ToolCallResponse Execute(ICSISapModelConnectionService service, string argumentsJson)
        {
            var result = service.GetModelStatistics();
            if (result.IsSuccess)
            {
                return Ok(result.Message, new { Count = result.Data.ShellCount });
            }
            return Fail(result.Message);
        }
    }
}
