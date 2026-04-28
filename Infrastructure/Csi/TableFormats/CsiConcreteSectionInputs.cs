namespace ExcelCSIToolBoxAddIn.Infrastructure.Csi
{
    public class CsiConcreteRectangleSectionInput
    {
        public string SectionName { get; set; }
        public string MaterialName { get; set; }
        public double H { get; set; }
        public double B { get; set; }
    }

    public class CsiConcreteCircleSectionInput
    {
        public string SectionName { get; set; }
        public string MaterialName { get; set; }
        public double D { get; set; }
    }
}
