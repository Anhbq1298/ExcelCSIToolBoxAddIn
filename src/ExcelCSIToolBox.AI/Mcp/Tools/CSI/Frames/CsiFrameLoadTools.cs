using System.Collections.Generic;
using ExcelCSIToolBox.AI.Mcp.Contracts;
using ExcelCSIToolBox.AI.Mcp.Tools.CSI.Base;
using ExcelCSIToolBox.Core.Abstractions.CSI;
using ExcelCSIToolBox.Core.Models.CSI;

namespace ExcelCSIToolBox.AI.Mcp.Tools.CSI.Frames
{
    public sealed class LoadsFrameAssignDistributedTool : CsiWriteToolBase<FrameDistributedLoadArgs>
    {
        public LoadsFrameAssignDistributedTool(ICsiModelCommandService commandService) : base(commandService) { }
        public override string Name => "loads.frame.assign_distributed";
        public override string Title => "Assign Frame Distributed Load";
        public override string Category => "Loads";
        public override string SubCategory => "Frame";
        public override string Description => "Assigns distributed load to frame objects. Medium risk and requires confirmation.";
        public override CsiMethodRiskLevel RiskLevel => CsiMethodRiskLevel.Medium;
        public override bool RequiresConfirmation => true;
        public override bool SupportsDryRun => true;

        protected override ToolCallResponse Execute(FrameDistributedLoadArgs args)
        {
            return args.DryRun
                ? Preview(CommandService.PreviewAssignFrameDistributedLoad(args.FrameNames, args.LoadPattern, args.Direction, args.Value1, args.Value2))
                : Result(CommandService.AssignFrameDistributedLoad(args.FrameNames, args.LoadPattern, args.Direction, args.Value1, args.Value2, args.Confirmed));
        }
    }

    public sealed class LoadsFrameAssignPointLoadTool : CsiWriteToolBase<FramePointLoadArgs>
    {
        public LoadsFrameAssignPointLoadTool(ICsiModelCommandService commandService) : base(commandService) { }
        public override string Name => "loads.frame.assign_point_load";
        public override string Title => "Assign Frame Point Load";
        public override string Category => "Loads";
        public override string SubCategory => "Frame";
        public override string Description => "Assigns point load to frame objects. Medium risk and requires confirmation.";
        public override CsiMethodRiskLevel RiskLevel => CsiMethodRiskLevel.Medium;
        public override bool RequiresConfirmation => true;
        public override bool SupportsDryRun => true;

        protected override ToolCallResponse Execute(FramePointLoadArgs args)
        {
            return args.DryRun
                ? Preview(CommandService.PreviewAssignFramePointLoad(args.FrameNames, args.LoadPattern, args.Direction, args.Distance, args.Value))
                : Result(CommandService.AssignFramePointLoad(args.FrameNames, args.LoadPattern, args.Direction, args.Distance, args.Value, args.Confirmed));
        }
    }

    public sealed class FramesAssignDistributedLoadTool : CsiWriteToolBase<FrameDistributedLoadArgs>
    {
        public FramesAssignDistributedLoadTool(ICsiModelCommandService commandService) : base(commandService) { }
        public override string Name => "frames.assign_distributed_load";
        public override string Title => "Assign Frame Distributed Load";
        public override string Category => "Frames";
        public override string SubCategory => "Loads";
        public override string Description => "Assigns distributed load to frame objects. Medium risk and requires confirmation.";
        public override CsiMethodRiskLevel RiskLevel => CsiMethodRiskLevel.Medium;
        public override bool RequiresConfirmation => true;
        public override bool SupportsDryRun => true;

        protected override ToolCallResponse Execute(FrameDistributedLoadArgs args)
        {
            return args.DryRun
                ? Preview(CommandService.PreviewAssignFrameDistributedLoad(args.FrameNames, args.LoadPattern, args.Direction, args.Value1, args.Value2))
                : Result(CommandService.AssignFrameDistributedLoad(args.FrameNames, args.LoadPattern, args.Direction, args.Value1, args.Value2, args.Confirmed));
        }
    }

    public sealed class FramesAssignPointLoadTool : CsiWriteToolBase<FramePointLoadArgs>
    {
        public FramesAssignPointLoadTool(ICsiModelCommandService commandService) : base(commandService) { }
        public override string Name => "frames.assign_point_load";
        public override string Title => "Assign Frame Point Load";
        public override string Category => "Frames";
        public override string SubCategory => "Loads";
        public override string Description => "Assigns point load to frame objects. Medium risk and requires confirmation.";
        public override CsiMethodRiskLevel RiskLevel => CsiMethodRiskLevel.Medium;
        public override bool RequiresConfirmation => true;
        public override bool SupportsDryRun => true;

        protected override ToolCallResponse Execute(FramePointLoadArgs args)
        {
            return args.DryRun
                ? Preview(CommandService.PreviewAssignFramePointLoad(args.FrameNames, args.LoadPattern, args.Direction, args.Distance, args.Value))
                : Result(CommandService.AssignFramePointLoad(args.FrameNames, args.LoadPattern, args.Direction, args.Distance, args.Value, args.Confirmed));
        }
    }

    public sealed class FrameDistributedLoadArgs : DryRunConfirmedArgs
    {
        public List<string> FrameNames { get; set; }
        public string LoadPattern { get; set; }
        public int Direction { get; set; }
        public double Value1 { get; set; }
        public double Value2 { get; set; }
    }

    public sealed class FramePointLoadArgs : DryRunConfirmedArgs
    {
        public List<string> FrameNames { get; set; }
        public string LoadPattern { get; set; }
        public int Direction { get; set; }
        public double Distance { get; set; }
        public double Value { get; set; }
    }
}
