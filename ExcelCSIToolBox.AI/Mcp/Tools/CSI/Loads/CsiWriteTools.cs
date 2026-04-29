using System.Collections.Generic;
using ExcelCSIToolBox.AI.Mcp.Tools.CSI.Base;
using ExcelCSIToolBox.AI.Mcp.Contracts;
using ExcelCSIToolBox.Core.Abstractions.CSI;
using ExcelCSIToolBox.Core.Models.CSI;

namespace ExcelCSIToolBox.AI.Mcp.Tools.CSI.Loads
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

    public sealed class SelectionClearTool : CsiWriteToolBase<DryRunConfirmedArgs>
    {
        public SelectionClearTool(ICsiModelCommandService commandService) : base(commandService) { }
        public override string Name => "selection.clear";
        public override string Title => "Clear Selection";
        public override string Category => "Selection";
        public override string SubCategory => "General";
        public override string Description => "Clears current CSI object selection. Low-risk write tool with dry-run support.";
        public override CsiMethodRiskLevel RiskLevel => CsiMethodRiskLevel.Low;
        public override bool RequiresConfirmation => false;
        public override bool SupportsDryRun => true;

        protected override ToolCallResponse Execute(DryRunConfirmedArgs args)
        {
            return args.DryRun
                ? Preview(CommandService.PreviewClearSelection())
                : Result(CommandService.ClearSelection(args.Confirmed));
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

    public sealed class AnalysisRunTool : CsiWriteToolBase<DryRunConfirmedArgs>
    {
        public AnalysisRunTool(ICsiModelCommandService commandService) : base(commandService) { }
        public override string Name => "analysis.run";
        public override string Title => "Run Analysis";
        public override string Category => "Analysis";
        public override string SubCategory => "Run";
        public override string Description => "Runs model analysis. High risk and requires explicit confirmation.";
        public override CsiMethodRiskLevel RiskLevel => CsiMethodRiskLevel.High;
        public override bool RequiresConfirmation => true;
        public override bool SupportsDryRun => false;

        protected override ToolCallResponse Execute(DryRunConfirmedArgs args)
        {
            return args.DryRun
                ? Preview(CommandService.PreviewRunAnalysis())
                : Result(CommandService.RunAnalysis(args.Confirmed));
        }
    }

    public sealed class FileSaveModelTool : CsiWriteToolBase<DryRunConfirmedArgs>
    {
        public FileSaveModelTool(ICsiModelCommandService commandService) : base(commandService) { }
        public override string Name => "file.save_model";
        public override string Title => "Save Model";
        public override string Category => "Model / File / Units";
        public override string SubCategory => "File";
        public override string Description => "Attempts to save the model. Dangerous and blocked by default for AI usage.";
        public override CsiMethodRiskLevel RiskLevel => CsiMethodRiskLevel.Dangerous;
        public override bool RequiresConfirmation => true;
        public override bool SupportsDryRun => false;

        protected override ToolCallResponse Execute(DryRunConfirmedArgs args)
        {
            return args.DryRun
                ? Preview(CommandService.PreviewSaveModel())
                : Result(CommandService.SaveModel(args.Confirmed));
        }
    }

    public sealed class PointByCoordinatesArgs : DryRunConfirmedArgs
    {
        public double X { get; set; }
        public double Y { get; set; }
        public double Z { get; set; }
        public string UserName { get; set; }
    }

    public sealed class FrameByCoordinatesArgs : DryRunConfirmedArgs
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

    public sealed class FrameByPointsArgs : DryRunConfirmedArgs
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

    public sealed class DeleteObjectsArgs : DryRunConfirmedArgs
    {
        public List<string> ObjectNames { get; set; }
    }

}
