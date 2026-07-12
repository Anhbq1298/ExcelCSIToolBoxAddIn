using System.Collections.Generic;

namespace ExcelCSIToolBox.Core.Contracts.CSI
{
    public class CSISapModelDisplayTableDTO
    {
        public IReadOnlyList<string> FieldKeys { get; set; }
        public IReadOnlyList<object[]> Rows { get; set; }
    }
}
