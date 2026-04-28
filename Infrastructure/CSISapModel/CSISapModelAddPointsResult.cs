using System.Collections.Generic;

namespace ExcelCSIToolBoxAddIn.Infrastructure.CSISapModel
{
    public class CSISapModelAddPointsResult
    {
        public int AddedCount { get; set; }

        public IReadOnlyList<string> FailedRowMessages { get; set; }
    }
}
