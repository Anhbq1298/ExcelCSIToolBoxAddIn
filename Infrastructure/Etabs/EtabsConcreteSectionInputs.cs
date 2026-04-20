namespace ExcelCSIToolBoxAddIn.Infrastructure.Etabs
{
    public class EtabsConcreteRectangleSectionInput
    {
        public string SectionName { get; set; }
        public string MaterialName { get; set; }
        public double H { get; set; }
        public double B { get; set; }
    }

    public class EtabsConcreteCircleSectionInput
    {
        public string SectionName { get; set; }
        public string MaterialName { get; set; }
        public double D { get; set; }
    }
}
