using System.Collections.Generic;

namespace ExcelCSIToolBox.Core.Contracts.CSI.PileCap
{
    public enum PileCapArrangementType
    {
        Mono = 0,
        TwoPile = 1,
        ThreePile = 2,
        FourPile = 3
    }

    public class PileCapAssignmentRequestDto
    {
        public PileCapArrangementType ArrangementType { get; set; }

        public double PileDiameterMillimeters { get; set; }

        public double PileLengthMillimeters { get; set; }

        public string PileMaterial { get; set; }

        public double RotationDegrees { get; set; }

        public bool AutoSpacing { get; set; }

        public double PileSpacingMillimeters { get; set; }

        public double SpacingXMillimeters { get; set; }

        public double SpacingYMillimeters { get; set; }

        public double PileCapThicknessMillimeters { get; set; }

        public double EdgeDistanceMillimeters { get; set; }

        public string PileCapMaterial { get; set; }

        public bool SelectCreatedObjects { get; set; }
    }

    public class PileCapAssignmentSummaryDto
    {
        public PileCapAssignmentSummaryDto()
        {
            CreatedFrameNames = new List<string>();
            CreatedAreaNames = new List<string>();
            SelectedPointNames = new List<string>();
            Warnings = new List<string>();
            Errors = new List<string>();
        }

        public int SelectedPointCount { get; set; }

        public int SuccessfullyProcessedPointCount { get; set; }

        public int CreatedPileCapCount { get; set; }

        public int CreatedPileCount { get; set; }

        public int SkippedPointCount { get; set; }

        public int FailedPointCount { get; set; }

        public int IgnoredNonPointObjectCount { get; set; }

        public string PilePropertyName { get; set; }

        public string PileCapPropertyName { get; set; }

        public IList<string> SelectedPointNames { get; private set; }

        public IList<string> CreatedFrameNames { get; private set; }

        public IList<string> CreatedAreaNames { get; private set; }

        public IList<string> Warnings { get; private set; }

        public IList<string> Errors { get; private set; }
    }
}
