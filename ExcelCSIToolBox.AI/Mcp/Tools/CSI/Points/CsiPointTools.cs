using System.Collections.Generic;
using ExcelCSIToolBox.AI.Mcp.Contracts;
using ExcelCSIToolBox.AI.Mcp.Tools.CSI.Base;
using ExcelCSIToolBox.Core.Abstractions.CSI;
using ExcelCSIToolBox.Core.Models.CSI;

namespace ExcelCSIToolBox.AI.Mcp.Tools.CSI.Points
{
    public sealed class PointsGetSelectedTool : CsiActiveConnectionToolBase
    {
        public PointsGetSelectedTool(ICSISapModelConnectionService etabsService, ICSISapModelConnectionService sap2000Service) : base(etabsService, sap2000Service) { }
        public override string Name => "points.get_selected";
        public override string Title => "Get Selected Points";
        public override string Category => "Points";
        public override string SubCategory => "Read";
        public override string Description => "Returns selected point objects with coordinates from the active CSI model.";
        public override bool IsReadOnly => true;
        public override CsiMethodRiskLevel RiskLevel => CsiMethodRiskLevel.None;
        public override bool RequiresConfirmation => false;
        public override bool SupportsDryRun => false;

        protected override ToolCallResponse Execute(ICSISapModelConnectionService service, string argumentsJson)
        {
            return Result(service.GetSelectedPointsFromActiveModel());
        }
    }

    public sealed class PointsSetSelectedTool : CsiActiveConnectionToolBase
    {
        public PointsSetSelectedTool(ICSISapModelConnectionService etabsService, ICSISapModelConnectionService sap2000Service) : base(etabsService, sap2000Service) { }
        public override string Name => "points.set_selected";
        public override string Title => "Select Points By Name";
        public override string Category => "Points";
        public override string SubCategory => "Selection";
        public override string Description => "Selects point objects by unique names. Low-risk selection write tool with dry-run support.";
        public override bool IsReadOnly => false;
        public override CsiMethodRiskLevel RiskLevel => CsiMethodRiskLevel.Low;
        public override bool RequiresConfirmation => false;
        public override bool SupportsDryRun => true;

        protected override ToolCallResponse Execute(ICSISapModelConnectionService service, string argumentsJson)
        {
            NamesDryRunArgs args = ReadArgs<NamesDryRunArgs>(argumentsJson);
            if (args.DryRun)
            {
                return Preview(new CsiWritePreview
                {
                    OperationName = Name,
                    RiskLevel = RiskLevel,
                    RequiresConfirmation = false,
                    SupportsDryRun = true,
                    Summary = $"This will select {Count(args.Names)} point object(s).",
                    AffectedObjects = args.Names ?? new List<string>()
                });
            }

            return Result(service.SelectPointsByUniqueNames(args.Names));
        }
    }
}
