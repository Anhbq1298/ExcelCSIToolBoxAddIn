using System;
using System.Threading;
using System.Threading.Tasks;
using ExcelCSIToolBox.AI.Mcp.Contracts;
using ExcelCSIToolBox.AI.Mcp.Tools.CSI.Base;
using ExcelCSIToolBox.Core.Abstractions.CSI;
using ExcelCSIToolBox.Core.Models.CSI;

namespace ExcelCSIToolBox.AI.Mcp.Tools.CSI.Model
{
    /// <summary>
    /// Read-only MCP tool: returns object counts for the active CSI model.
    /// This is highly useful for answering questions like "how many points are there?".
    /// </summary>
    public sealed class CsiGetModelStatisticsTool : CsiActiveConnectionToolBase
    {
        public CsiGetModelStatisticsTool(ICSISapModelConnectionService etabsService, ICSISapModelConnectionService sap2000Service) 
            : base(etabsService, sap2000Service) { }

        public override string Name => "csi.get_model_statistics";
        public override string Title => "Get Model Statistics";
        public override string Category => "Model";
        public override string SubCategory => "Read";
        public override string Description => "Returns counts of points, frames, shells, load patterns, and combinations in the active CSI model. Use this to answer questions about model size or object counts.";
        public override bool IsReadOnly => true;
        public override CsiMethodRiskLevel RiskLevel => CsiMethodRiskLevel.None;
        public override bool RequiresConfirmation => false;
        public override bool SupportsDryRun => false;

        protected override ToolCallResponse Execute(ICSISapModelConnectionService service, string argumentsJson)
        {
            return Result(service.GetModelStatistics());
        }
    }
}
