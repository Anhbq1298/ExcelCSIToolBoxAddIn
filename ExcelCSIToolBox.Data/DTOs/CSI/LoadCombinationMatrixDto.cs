using System.Collections.Generic;

namespace ExcelCSIToolBox.Data.DTOs.CSI
{
    public class LoadCombinationMatrixDto
    {
        public List<string> LoadPatternNames { get; set; } = new List<string>();
        public List<LoadCombinationMatrixRowDto> Rows { get; set; } = new List<LoadCombinationMatrixRowDto>();
    }
}
