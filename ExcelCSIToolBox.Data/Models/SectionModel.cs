using ExcelCSIToolBox.Data.Models;
namespace ExcelCSIToolBox.Data.Models
{
    public class CSISapModelSteelISectionInput
    {
        public string SectionName { get; set; }
        public string MaterialName { get; set; }
        public double H { get; set; }
        public double B { get; set; }
        public double Tw { get; set; }
        public double Tf { get; set; }
    }

    public class CSISapModelSteelChannelSectionInput
    {
        public string SectionName { get; set; }
        public string MaterialName { get; set; }
        public double H { get; set; }
        public double B { get; set; }
        public double Tw { get; set; }
        public double Tf { get; set; }
    }

    public class CSISapModelSteelAngleSectionInput
    {
        public string SectionName { get; set; }
        public string MaterialName { get; set; }
        public double H { get; set; }
        public double B { get; set; }
        public double Tw { get; set; }
        public double Tf { get; set; }
    }

    public class CSISapModelSteelPipeSectionInput
    {
        public string SectionName { get; set; }
        public string MaterialName { get; set; }
        public double OutsideDiameter { get; set; }
        public double WallThickness { get; set; }
    }

    public class CSISapModelSteelTubeSectionInput
    {
        public string SectionName { get; set; }
        public string MaterialName { get; set; }
        public double H { get; set; }
        public double B { get; set; }
        public double T { get; set; }
    }

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


