using System.Collections.Generic;

namespace ExcelCSIToolBox.Data.DTOs.CSI
{
    public class CSISapModelDisplayTableDTO
    {
        public IReadOnlyList<string> FieldKeys { get; set; }
        public IReadOnlyList<object[]> Rows { get; set; }
    }
}
