namespace ExcelCSIToolBoxAddIn.Core.Tabular
{
    public class ExcelConcreteRectangleSectionRow
    {
        public int ExcelRowNumber { get; set; }
        public string SectionName { get; set; }
        public string MaterialName { get; set; }
        public string HText { get; set; }
        public string BText { get; set; }
    }

    public class ExcelConcreteCircleSectionRow
    {
        public int ExcelRowNumber { get; set; }
        public string SectionName { get; set; }
        public string MaterialName { get; set; }
        public string DText { get; set; }
    }
}
