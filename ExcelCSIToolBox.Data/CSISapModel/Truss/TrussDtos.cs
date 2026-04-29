using System.Collections.Generic;

namespace ExcelCSIToolBox.Data.CSISapModel.Truss
{
    public sealed class HoweTrussRequestDto
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
    }

    public sealed class HoweTrussResultDto
    {
        public bool Success { get; set; }
        public int BayCount { get; set; }
        public double Span { get; set; }
        public double BaySpacing { get; set; }
        public int AddedFrameCount { get; set; }
        public int ReleasedWebMemberCount { get; set; }
        public List<string> ChordFrameNames { get; set; }
        public List<string> WebFrameNames { get; set; }
        public List<string> FailureReasons { get; set; }
    }
}
