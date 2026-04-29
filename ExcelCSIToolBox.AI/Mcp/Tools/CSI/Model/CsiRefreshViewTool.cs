using System;
using Newtonsoft.Json;
using ExcelCSIToolBox.AI.Mcp.Contracts;
using ExcelCSIToolBox.AI.Mcp.Tools.CSI.Base;
using ExcelCSIToolBox.Core.Abstractions.CSI;
using ExcelCSIToolBox.Core.Models.CSI;

namespace ExcelCSIToolBox.AI.Mcp.Tools.CSI.Model
{
    /// <summary>
    /// MCP tool to refresh the model view in ETABS or SAP2000.
    /// Supports an optional zoomAll parameter.
    /// </summary>
    public sealed class CsiRefreshViewTool : CsiActiveConnectionToolBase
    {
        public CsiRefreshViewTool(ICSISapModelConnectionService etabsService, ICSISapModelConnectionService sap2000Service) 
            : base(etabsService, sap2000Service) { }

        public override string Name => "csi.refresh_view";
        public override string Title => "Refresh Model View";
        public override string Category => "Model / View";
        public override string SubCategory => "Action";
        public override string Description => "Refreshes the active CSI model view. Use this if objects were added but are not visible. Set zoomAll to true to fit all objects in view.";
        public override bool IsReadOnly => false; // Modifies view state
        public override CsiMethodRiskLevel RiskLevel => CsiMethodRiskLevel.None;
        public override bool RequiresConfirmation => false;
        public override bool SupportsDryRun => false;

        protected override ToolCallResponse Execute(ICSISapModelConnectionService service, string argumentsJson)
        {
            var args = JsonConvert.DeserializeObject<RefreshViewArgs>(argumentsJson) ?? new RefreshViewArgs();
            var result = service.RefreshView(args.ZoomAll);
            
            if (result.IsSuccess)
            {
                return Ok(args.ZoomAll ? "View refreshed and zoomed to fit all objects." : "View refreshed.");
            }
            return Fail(result.Message);
        }

        private class RefreshViewArgs
        {
            [JsonProperty("zoomAll")]
            public bool ZoomAll { get; set; } = false;
        }
    }
}
