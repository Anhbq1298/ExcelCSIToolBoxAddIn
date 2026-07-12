namespace ExcelCSIToolBox.Core.Contracts.CSI
{
    public class CSISapModelOutputCaseDTO
    {
        public string Name { get; set; }

        public string Type { get; set; }

        public bool IsLoadCombination { get; set; }

        public bool IsSeismicWindOrResponseSpectrum { get; set; }

        public string Kind
        {
            get { return IsLoadCombination ? "Load Combination" : "Load Case"; }
        }
    }
}
