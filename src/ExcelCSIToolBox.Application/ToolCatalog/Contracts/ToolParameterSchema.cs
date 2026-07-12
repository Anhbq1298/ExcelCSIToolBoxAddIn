namespace ExcelCSIToolBox.Application.ToolCatalog.Contracts
{
    public sealed class ToolParameterSchema
    {
        public string Name { get; set; }
        public bool IsRequired { get; set; }
        public string ValueType { get; set; }
    }
}
