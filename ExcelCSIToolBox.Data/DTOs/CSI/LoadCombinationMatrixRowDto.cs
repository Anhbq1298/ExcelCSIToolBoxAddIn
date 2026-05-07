using System.Collections.Generic;

namespace ExcelCSIToolBox.Data.DTOs.CSI
{
    public class LoadCombinationMatrixRowDto
    {
        public string LoadCombinationName { get; set; }
        public int CombinationType { get; set; }
        public Dictionary<string, double?> Factors { get; set; } = new Dictionary<string, double?>();
        public Dictionary<string, int> FactorCaseTypes { get; set; } = new Dictionary<string, int>();
    }
}
