namespace ExcelCSIToolBoxAddIn.Infrastructure.CSISapModel
{
    public class CSISapModelConcreteRectangleSectionInput
    {
        public string SectionName { get; set; }
        public string MaterialName { get; set; }
        public double H { get; set; }
        public double B { get; set; }
    }

    public class CSISapModelConcreteCircleSectionInput
    {
        public string SectionName { get; set; }
        public string MaterialName { get; set; }
        public double D { get; set; }
    }
}
