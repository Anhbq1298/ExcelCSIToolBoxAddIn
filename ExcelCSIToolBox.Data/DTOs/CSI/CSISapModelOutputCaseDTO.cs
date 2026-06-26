namespace ExcelCSIToolBox.Data.DTOs.CSI
{
    public class CSISapModelOutputCaseDTO
    {
        public string Name { get; set; }

        public string Type { get; set; }

        public bool IsLoadCombination { get; set; }

        public string Kind
        {
            get { return IsLoadCombination ? "Load Combination" : "Load Case"; }
        }
    }
}
