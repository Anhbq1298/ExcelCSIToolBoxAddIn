namespace ExcelCSIToolBoxAddIn.Infrastructure.Etabs
{
    public class EtabsFrameByPointInput
    {
        public int ExcelRowNumber { get; set; }

        public string UniqueName { get; set; }

        public string Section { get; set; }

        public string PointIName { get; set; }

        public string PointJName { get; set; }
    }
}
