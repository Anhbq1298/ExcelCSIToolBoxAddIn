using System;
using System.Collections.Generic;
using ExcelCSIToolBox.AI.Mcp.Contracts;
using ExcelCSIToolBox.AI.Mcp.Tools.CSI.Base;
using ExcelCSIToolBox.Core.Abstractions.CSI;
using ExcelCSIToolBox.Core.Models.CSI;

namespace ExcelCSIToolBox.AI.Mcp.Tools.CSI.Frames
{
    public sealed class AnalyzeSelectedFramesTool : CsiActiveConnectionToolBase
    {
        public AnalyzeSelectedFramesTool(ICSISapModelConnectionService etabsService, ICSISapModelConnectionService sap2000Service) 
            : base(etabsService, sap2000Service) { }

        public override string Name => "frames.analyze_selected";
        public override string Title => "Analyze Selected Frames";
        public override string Category => "Frames";
        public override string SubCategory => "Analysis";
        public override string Description => "Collects detailed info (geometry, section, loads) for all selected frames in a single call.";
        public override bool IsReadOnly => true;
        public override CsiMethodRiskLevel RiskLevel => CsiMethodRiskLevel.None;
        public override bool RequiresConfirmation => false;
        public override bool SupportsDryRun => false;

        protected override ToolCallResponse Execute(ICSISapModelConnectionService service, string argumentsJson)
        {
            var selectedNamesResult = service.GetSelectedFramesFromActiveModel();
            if (!selectedNamesResult.IsSuccess) return Fail(selectedNamesResult.Message);

            var frameList = new List<object>();
            foreach (var name in selectedNamesResult.Data)
            {
                var section = service.GetFrameSection(name).Data;
                var points = service.GetFramePoints(name).Data;
                var distLoads = service.GetFrameDistributedLoads(name).Data;

                frameList.Add(new
                {
                    Name = name,
                    Section = section?.SectionName,
                    PointI = points?.PointI,
                    PointJ = points?.PointJ,
                    DistributedLoads = distLoads
                });
            }

            return Ok($"Collected info for {frameList.Count} selected frame(s).", new { Frames = frameList });
        }
    }
}
