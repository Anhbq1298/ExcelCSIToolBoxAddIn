namespace ExcelCSIToolBoxAddIn.Infrastructure.Etabs
{
    public class EtabsSteelISectionInput
    {
        public string SectionName { get; set; }
        public string MaterialName { get; set; }
        public double H { get; set; }
        public double B { get; set; }
        public double Tw { get; set; }
        public double Tf { get; set; }
    }

    public class EtabsSteelPipeSectionInput
    {
        public string SectionName { get; set; }
        public string MaterialName { get; set; }
        public double OutsideDiameter { get; set; }
        public double WallThickness { get; set; }
    }

    public class EtabsSteelTubeSectionInput
    {
        public string SectionName { get; set; }
        public string MaterialName { get; set; }
        public double H { get; set; }
        public double B { get; set; }
        public double T { get; set; }
    }
}
