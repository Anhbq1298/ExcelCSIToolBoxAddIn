using System;

namespace ExcelCSIToolBoxAddIn.Core.Tabular
{
    public class ExcelSteelISectionRow
    {
        public int ExcelRowNumber { get; set; }
        public string SectionName { get; set; }
        public string MaterialName { get; set; }
        public string HText { get; set; }
        public string BText { get; set; }
        public string TwText { get; set; }
        public string TfText { get; set; }
    }

    public class ExcelSteelChannelSectionRow
    {
        public int ExcelRowNumber { get; set; }
        public string SectionName { get; set; }
        public string MaterialName { get; set; }
        public string HText { get; set; }
        public string BText { get; set; }
        public string TwText { get; set; }
        public string TfText { get; set; }
    }

    public class ExcelSteelAngleSectionRow
    {
        public int ExcelRowNumber { get; set; }
        public string SectionName { get; set; }
        public string MaterialName { get; set; }
        public string HText { get; set; }
        public string BText { get; set; }
        public string TwText { get; set; }
        public string TfText { get; set; }
    }

    public class ExcelSteelPipeSectionRow
    {
        public int ExcelRowNumber { get; set; }
        public string SectionName { get; set; }
        public string MaterialName { get; set; }
        public string OutsideDiameterText { get; set; }
        public string WallThicknessText { get; set; }
    }

    public class ExcelSteelTubeSectionRow
    {
        public int ExcelRowNumber { get; set; }
        public string SectionName { get; set; }
        public string MaterialName { get; set; }
        public string HText { get; set; }
        public string BText { get; set; }
        public string TText { get; set; }
    }
}
