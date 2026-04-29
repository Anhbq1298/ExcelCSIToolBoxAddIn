using ExcelCSIToolBox.AI.Mcp.Contracts;
using ExcelCSIToolBox.AI.Mcp.Tools.CSI.Base;
using ExcelCSIToolBox.Core.Abstractions.CSI;
using ExcelCSIToolBox.Core.Models.CSI;

namespace ExcelCSIToolBox.AI.Mcp.Tools.CSI.Points
{
    public sealed class PointsAddByCoordinatesTool : CsiWriteToolBase<PointByCoordinatesArgs>
    {
        public PointsAddByCoordinatesTool(ICsiModelCommandService commandService) : base(commandService) { }
        public override string Name => "points.add_by_coordinates";
        public override string Title => "Add Point By Coordinates";
        public override string Category => "Points";
        public override string SubCategory => "Creation";
        public override string Description => "Adds one CSI point by Cartesian coordinates. Low-risk write tool with dry-run support.";
        public override CsiMethodRiskLevel RiskLevel => CsiMethodRiskLevel.Low;
        public override bool RequiresConfirmation => false;
        public override bool SupportsDryRun => true;

        protected override ToolCallResponse Execute(PointByCoordinatesArgs args)
        {
            return args.DryRun
                ? Preview(CommandService.PreviewAddPoint(args.X, args.Y, args.Z, args.UserName))
                : Result(CommandService.AddPoint(args.X, args.Y, args.Z, args.UserName, args.Confirmed));
        }
    }

    public sealed class PointByCoordinatesArgs : DryRunConfirmedArgs
    {
        public double X { get; set; }
        public double Y { get; set; }
        public double Z { get; set; }
        public string UserName { get; set; }
    }
}
