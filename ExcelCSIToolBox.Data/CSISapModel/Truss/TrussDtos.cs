using System.Collections.Generic;

namespace ExcelCSIToolBox.Data.CSISapModel.Truss
{
    public class HoweTrussRequestDto
    {
        public int BayCount { get; set; }
        public double Span { get; set; }
        public double Height { get; set; }
        public double StartX { get; set; }
        public double StartY { get; set; }
        public double StartZ { get; set; }
        public string NamePrefix { get; set; }
        public string ChordPropName { get; set; }
        public string WebPropName { get; set; }
        public string TrussType { get; set; }
        public string SlopeMode { get; set; }
        public double Slope { get; set; }
        public string MonoSlopeDirection { get; set; }
        public string DistributedLoadPattern { get; set; }
        public int DistributedLoadDirection { get; set; }
        public double DistributedLoadValue1 { get; set; }
        public double DistributedLoadValue2 { get; set; }
        public string DistributedLoadTarget { get; set; }
    }

    public class HoweTrussResultDto
    {
        public bool Success { get; set; }
        public int BayCount { get; set; }
        public double Span { get; set; }
        public double BaySpacing { get; set; }
        public string TrussType { get; set; }
        public string SlopeMode { get; set; }
        public double Slope { get; set; }
        public string MonoSlopeDirection { get; set; }
        public string DistributedLoadPattern { get; set; }
        public int DistributedLoadDirection { get; set; }
        public double DistributedLoadValue1 { get; set; }
        public double DistributedLoadValue2 { get; set; }
        public string DistributedLoadTarget { get; set; }
        public int LoadedFrameCount { get; set; }
        public int AddedFrameCount { get; set; }
        public int ReleasedWebMemberCount { get; set; }
        public List<string> ChordFrameNames { get; set; }
        public List<string> WebFrameNames { get; set; }
        public List<string> FailureReasons { get; set; }
    }

    public sealed class PrattTrussRequestDto : HoweTrussRequestDto
    {
    }

    public sealed class PrattTrussResultDto : HoweTrussResultDto
    {
    }
}
