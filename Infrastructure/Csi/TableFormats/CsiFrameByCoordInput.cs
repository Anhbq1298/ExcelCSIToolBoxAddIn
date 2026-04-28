namespace ExcelCSIToolBoxAddIn.Infrastructure.Csi
{
    public class CsiFrameByCoordInput
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
}
