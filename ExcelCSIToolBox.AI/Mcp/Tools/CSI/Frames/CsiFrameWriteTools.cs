using System.Collections.Generic;
using ExcelCSIToolBox.AI.Mcp.Contracts;
using ExcelCSIToolBox.AI.Mcp.Tools.CSI.Base;
using ExcelCSIToolBox.Core.Abstractions.CSI;
using ExcelCSIToolBox.Core.Models.CSI;

namespace ExcelCSIToolBox.AI.Mcp.Tools.CSI.Frames
{
    public sealed class FramesAddByCoordinatesTool : CsiWriteToolBase<FrameByCoordinatesArgs>
    {
        public FramesAddByCoordinatesTool(ICsiModelCommandService commandService) : base(commandService) { }
        public override string Name => "frames.add_by_coordinates";
        public override string Title => "Add Frame By Coordinates";
        public override string Category => "Frames";
        public override string SubCategory => "Creation";
        public override string Description => "Adds one frame by end coordinates. Low-risk write tool with dry-run support.";
        public override CsiMethodRiskLevel RiskLevel => CsiMethodRiskLevel.Low;
        public override bool RequiresConfirmation => false;
        public override bool SupportsDryRun => true;

        protected override ToolCallResponse Execute(FrameByCoordinatesArgs args)
        {
            return args.DryRun
                ? Preview(CommandService.PreviewAddFrameByCoordinates(args.Xi, args.Yi, args.Zi, args.Xj, args.Yj, args.Zj, args.SectionName, args.UserName))
                : Result(CommandService.AddFrameByCoordinates(args.Xi, args.Yi, args.Zi, args.Xj, args.Yj, args.Zj, args.SectionName, args.UserName, args.Confirmed));
        }
    }

    public sealed class FramesAddByPointsTool : CsiWriteToolBase<FrameByPointsArgs>
    {
        public FramesAddByPointsTool(ICsiModelCommandService commandService) : base(commandService) { }
        public override string Name => "frames.add_by_points";
        public override string Title => "Add Frame By Points";
        public override string Category => "Frames";
        public override string SubCategory => "Creation";
        public override string Description => "Adds one frame between two existing point objects. Low-risk write tool with dry-run support.";
        public override CsiMethodRiskLevel RiskLevel => CsiMethodRiskLevel.Low;
        public override bool RequiresConfirmation => false;
        public override bool SupportsDryRun => true;

        protected override ToolCallResponse Execute(FrameByPointsArgs args)
        {
            return args.DryRun
                ? Preview(CommandService.PreviewAddFrameByPoints(args.Point1Name, args.Point2Name, args.SectionName, args.UserName))
                : Result(CommandService.AddFrameByPoints(args.Point1Name, args.Point2Name, args.SectionName, args.UserName, args.Confirmed));
        }
    }

    public sealed class FramesAssignSectionTool : CsiWriteToolBase<FramesAssignSectionArgs>
    {
        public FramesAssignSectionTool(ICsiModelCommandService commandService) : base(commandService) { }
        public override string Name => "frames.assign_section";
        public override string Title => "Assign Frame Section";
        public override string Category => "Frames";
        public override string SubCategory => "Assignments";
        public override string Description => "Assigns a section property to one or more frames. Medium risk and requires confirmation.";
        public override CsiMethodRiskLevel RiskLevel => CsiMethodRiskLevel.Medium;
        public override bool RequiresConfirmation => true;
        public override bool SupportsDryRun => true;

        protected override ToolCallResponse Execute(FramesAssignSectionArgs args)
        {
            return args.DryRun
                ? Preview(CommandService.PreviewAssignFrameSection(args.FrameNames, args.SectionName))
                : Result(CommandService.AssignFrameSection(args.FrameNames, args.SectionName, args.Confirmed));
        }
    }

    public sealed class FramesDeleteTool : CsiWriteToolBase<DeleteObjectsArgs>
    {
        public FramesDeleteTool(ICsiModelCommandService commandService) : base(commandService) { }
        public override string Name => "frames.delete";
        public override string Title => "Delete Frames";
        public override string Category => "Frames";
        public override string SubCategory => "Deletion";
        public override string Description => "Deletes frame objects. High risk and requires explicit confirmation.";
        public override CsiMethodRiskLevel RiskLevel => CsiMethodRiskLevel.High;
        public override bool RequiresConfirmation => true;
        public override bool SupportsDryRun => true;

        protected override ToolCallResponse Execute(DeleteObjectsArgs args)
        {
            return args.DryRun
                ? Preview(CommandService.PreviewDeleteObjects(args.ObjectNames, "frames"))
                : Result(CommandService.DeleteObjects(args.ObjectNames, "frames", args.Confirmed));
        }
    }

    public sealed class FrameByCoordinatesArgs : LowRiskWriteArgs
    {
        public double Xi { get; set; }
        public double Yi { get; set; }
        public double Zi { get; set; }
        public double Xj { get; set; }
        public double Yj { get; set; }
        public double Zj { get; set; }
        public string SectionName { get; set; }
        public string UserName { get; set; }
    }

    public sealed class FrameByPointsArgs : LowRiskWriteArgs
    {
        public string Point1Name { get; set; }
        public string Point2Name { get; set; }
        public string SectionName { get; set; }
        public string UserName { get; set; }
    }

    public sealed class FramesAssignSectionArgs : DryRunConfirmedArgs
    {
        public List<string> FrameNames { get; set; }
        public string SectionName { get; set; }
    }

    public sealed class DeleteObjectsArgs : DryRunConfirmedArgs
    {
        public List<string> ObjectNames { get; set; }
    }
}
