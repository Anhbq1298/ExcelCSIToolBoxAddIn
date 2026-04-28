using System.Collections.Generic;

namespace ExcelCSIToolBoxAddIn.Infrastructure.CSISapModel
{
    public class CSISapModelAddFramesResult
    {
        public int AddedCount { get; set; }

        public IReadOnlyList<string> FailedRowMessages { get; set; }
    }
}
