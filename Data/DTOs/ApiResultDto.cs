using ExcelCSIToolBoxAddIn.Data.DTOs;
using System.Collections.Generic;

namespace ExcelCSIToolBoxAddIn.Data.DTOs
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
        public object CsiObject { get; set; }
        public object SapModel { get; set; }
    }
}

