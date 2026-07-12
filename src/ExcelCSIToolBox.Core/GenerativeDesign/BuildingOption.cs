namespace ExcelCSIToolBox.Core.GenerativeDesign
{
    public sealed class BuildingOption
    {
        public string OptionId { get; set; }
        public string Name { get; set; }
        public StructuralScheme Scheme { get; set; }
        public double EstimatedWeight { get; set; }
        public double EstimatedCost { get; set; }
        public string Notes { get; set; }
    }
}
