using System.Collections.Generic;

namespace ExcelCSIToolBoxAddIn.Infrastructure.Etabs
{
    public class EtabsAddPointsResult
    {
        public int AddedCount { get; set; }

        public IReadOnlyList<int> FailedRows { get; set; }
    }
}
