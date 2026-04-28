namespace ExcelCSIToolBoxAddIn.Data.Models
{
    public class CSISapModelFrameByCoordInput
    {
        public int ExcelRowNumber { get; set; }
        public string UniqueName { get; set; }
        public string SectionName { get; set; }
        public double Xi { get; set; }
        public double Yi { get; set; }
        public double Zi { get; set; }
        public double Xj { get; set; }
        public double Yj { get; set; }
        public double Zj { get; set; }
    }

    public class CSISapModelFrameByPointInput
    {
        public int ExcelRowNumber { get; set; }
        public string UniqueName { get; set; }
        public string SectionName { get; set; }
        public string Point1Name { get; set; }
        public string Point2Name { get; set; }
    }
}
