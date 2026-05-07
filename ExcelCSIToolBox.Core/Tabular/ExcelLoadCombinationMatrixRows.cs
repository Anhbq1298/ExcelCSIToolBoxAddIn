using System.Collections.Generic;

namespace ExcelCSIToolBox.Core.Tabular
{
    /// <summary>
    /// Represents a load combination matrix read from an Excel range.
    /// </summary>
    public class ExcelLoadCombinationMatrix
    {
        public List<string> LoadPatternNames { get; set; } = new List<string>();

        public List<string> LoadCombinationReferenceNames { get; set; } = new List<string>();

        public List<ExcelLoadCombinationMatrixRow> Rows { get; set; } = new List<ExcelLoadCombinationMatrixRow>();
    }

    /// <summary>
    /// Represents a single load combination row read from an Excel matrix.
    /// </summary>
    public class ExcelLoadCombinationMatrixRow
    {
        public string LoadCombinationName { get; set; }

        public int CombinationType { get; set; }

        public Dictionary<string, double?> Factors { get; set; } = new Dictionary<string, double?>();

        public Dictionary<string, int> FactorCaseTypes { get; set; } = new Dictionary<string, int>();

        public Dictionary<string, double?> LoadCaseFactors { get; set; } = new Dictionary<string, double?>();

        public Dictionary<string, double?> LoadCombinationFactors { get; set; } = new Dictionary<string, double?>();
    }
}
