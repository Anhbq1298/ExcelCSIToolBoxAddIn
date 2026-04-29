using ExcelCSIToolBox.AI.Mcp.Contracts;
using ExcelCSIToolBox.AI.Mcp.Tools.CSI.Base;
using ExcelCSIToolBox.Core.Abstractions.CSI;
using ExcelCSIToolBox.Core.Models.CSI;

namespace ExcelCSIToolBox.AI.Mcp.Tools.CSI.Model
{
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
}
