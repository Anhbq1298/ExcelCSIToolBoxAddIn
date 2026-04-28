using System.Collections.Generic;

namespace ExcelCSIToolBoxAddIn.Infrastructure.Csi
{
    public class CsiAddFramesResult
    {
        public int AddedCount { get; set; }

        public IReadOnlyList<string> FailedRowMessages { get; set; }
    }
}
