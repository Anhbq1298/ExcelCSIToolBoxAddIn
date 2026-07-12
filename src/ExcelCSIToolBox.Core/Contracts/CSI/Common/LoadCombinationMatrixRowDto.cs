using System.Collections.Generic;

namespace ExcelCSIToolBox.Core.Contracts.CSI
{
    public class LoadCombinationMatrixRowDto
    {
        public string LoadCombinationName { get; set; }
        public int CombinationType { get; set; }
        public Dictionary<string, int> FactorCaseTypes { get; set; } = new Dictionary<string, int>();
        public Dictionary<string, double?> LoadCaseFactors { get; set; } = new Dictionary<string, double?>();
        public Dictionary<string, double?> LoadCombinationFactors { get; set; } = new Dictionary<string, double?>();
    }
}
