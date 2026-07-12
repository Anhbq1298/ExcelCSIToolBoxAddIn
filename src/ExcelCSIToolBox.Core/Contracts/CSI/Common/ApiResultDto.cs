using ExcelCSIToolBox.Core.Contracts.CSI;
using System.Collections.Generic;

namespace ExcelCSIToolBox.Core.Contracts.CSI
{
    public class CSISapModelAddFramesResultDTO
    {
        public int AddedCount { get; set; }
        public IReadOnlyList<string> FailedRowMessages { get; set; }
    }

    public class CSISapModelAddPointsResultDTO
    {
        public int AddedCount { get; set; }
        public IReadOnlyList<string> FailedRowMessages { get; set; }
    }

    public class CSISapModelConnectionInfoDTO
    {
        public bool IsConnected { get; set; }
        public string ModelPath { get; set; }
        public string ModelFileName { get; set; }
        public string ModelCurrentUnit { get; set; }
        public int? ProcessId { get; set; }
        public object CsiObject { get; set; }
        public object SapModel { get; set; }
    }

    public class CSISapModelStatisticsDTO
    {
        public int PointCount { get; set; }
        public int FrameCount { get; set; }
        public int ShellCount { get; set; }
        public int LoadPatternCount { get; set; }
        public int LoadCombinationCount { get; set; }
    }
}


