namespace ExcelCSIToolBoxAddIn.Infrastructure.Csi
{
    public class CsiSteelISectionInput
    {
        public string SectionName { get; set; }
        public string MaterialName { get; set; }
        public double H { get; set; }
        public double B { get; set; }
        public double Tw { get; set; }
        public double Tf { get; set; }
    }

    public class CsiSteelChannelSectionInput
    {
        public string SectionName { get; set; }
        public string MaterialName { get; set; }
        public double H { get; set; }
        public double B { get; set; }
        public double Tw { get; set; }
        public double Tf { get; set; }
    }

    public class CsiSteelAngleSectionInput
    {
        public string SectionName { get; set; }
        public string MaterialName { get; set; }
        public double H { get; set; }
        public double B { get; set; }
        public double Tw { get; set; }
        public double Tf { get; set; }
    }

    public class CsiSteelPipeSectionInput
    {
        public string SectionName { get; set; }
        public string MaterialName { get; set; }
        public double OutsideDiameter { get; set; }
        public double WallThickness { get; set; }
    }

    public class CsiSteelTubeSectionInput
    {
        public string SectionName { get; set; }
        public string MaterialName { get; set; }
        public double H { get; set; }
        public double B { get; set; }
        public double T { get; set; }
    }
}
