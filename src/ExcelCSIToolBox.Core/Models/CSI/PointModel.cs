using ExcelCSIToolBox.Core.Models.CSI;
namespace ExcelCSIToolBox.Core.Models.CSI
{
    public class CSISapModelPointCartesianInput
    {
        public int ExcelRowNumber { get; set; }
        public string UniqueName { get; set; }
        public double X { get; set; }
        public double Y { get; set; }
        public double Z { get; set; }
    }
}


