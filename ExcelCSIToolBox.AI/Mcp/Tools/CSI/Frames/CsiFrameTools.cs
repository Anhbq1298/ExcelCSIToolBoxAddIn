using System.Collections.Generic;
using ExcelCSIToolBox.AI.Mcp.Contracts;
using ExcelCSIToolBox.AI.Mcp.Tools.CSI.Base;
using ExcelCSIToolBox.Core.Abstractions.CSI;
using ExcelCSIToolBox.Core.Models.CSI;

namespace ExcelCSIToolBox.AI.Mcp.Tools.CSI.Frames
{
    public sealed class FramesGetSelectedTool : CsiActiveConnectionToolBase
    {
        public FramesGetSelectedTool(ICSISapModelConnectionService etabsService, ICSISapModelConnectionService sap2000Service) : base(etabsService, sap2000Service) { }
        public override string Name => "frames.get_selected";
        public override string Title => "Get Selected Frames";
        public override string Category => "Frames";
        public override string SubCategory => "Read";
        public override string Description => "Returns selected frame object names from the active CSI model.";
        public override bool IsReadOnly => true;
        public override CsiMethodRiskLevel RiskLevel => CsiMethodRiskLevel.None;
        public override bool RequiresConfirmation => false;
        public override bool SupportsDryRun => false;

        protected override ToolCallResponse Execute(ICSISapModelConnectionService service, string argumentsJson)
        {
            return Result(service.GetSelectedFramesFromActiveModel());
        }
    }

    public sealed class FramesSetSelectedTool : CsiActiveConnectionToolBase
    {
        public FramesSetSelectedTool(ICSISapModelConnectionService etabsService, ICSISapModelConnectionService sap2000Service) : base(etabsService, sap2000Service) { }
        public override string Name => "frames.set_selected";
        public override string Title => "Select Frames By Name";
        public override string Category => "Frames";
        public override string SubCategory => "Selection";
        public override string Description => "Selects frame objects by unique names. Low-risk selection write tool with dry-run support.";
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
                    Summary = $"This will select {Count(args.Names)} frame object(s).",
                    AffectedObjects = args.Names ?? new List<string>()
                });
            }

            return Result(service.SelectFramesByUniqueNames(args.Names));
        }
    }

    public sealed class FramesGetSectionsTool : CsiActiveConnectionToolBase
    {
        public FramesGetSectionsTool(ICSISapModelConnectionService etabsService, ICSISapModelConnectionService sap2000Service) : base(etabsService, sap2000Service) { }
        public override string Name => "frames.get_sections";
        public override string Title => "Get Frame Sections";
        public override string Category => "Frames";
        public override string SubCategory => "Properties";
        public override string Description => "Returns frame section property names and shape types from the active CSI model.";
        public override bool IsReadOnly => true;
        public override CsiMethodRiskLevel RiskLevel => CsiMethodRiskLevel.None;
        public override bool RequiresConfirmation => false;
        public override bool SupportsDryRun => false;

        protected override ToolCallResponse Execute(ICSISapModelConnectionService service, string argumentsJson)
        {
            return Result(service.GetFrameSections());
        }
    }

    public sealed class FramesGetSectionDetailTool : CsiActiveConnectionToolBase
    {
        public FramesGetSectionDetailTool(ICSISapModelConnectionService etabsService, ICSISapModelConnectionService sap2000Service) : base(etabsService, sap2000Service) { }
        public override string Name => "frames.get_section_detail";
        public override string Title => "Get Frame Section Detail";
        public override string Category => "Frames";
        public override string SubCategory => "Properties";
        public override string Description => "Returns detailed dimensions and material for one frame section property.";
        public override bool IsReadOnly => true;
        public override CsiMethodRiskLevel RiskLevel => CsiMethodRiskLevel.None;
        public override bool RequiresConfirmation => false;
        public override bool SupportsDryRun => false;

        protected override ToolCallResponse Execute(ICSISapModelConnectionService service, string argumentsJson)
        {
            SectionNameArgs args = ReadArgs<SectionNameArgs>(argumentsJson);
            return Result(service.GetFrameSectionDetail(args.SectionName));
        }
    }

    public sealed class SectionNameArgs
    {
        public string SectionName { get; set; }
    }
}
