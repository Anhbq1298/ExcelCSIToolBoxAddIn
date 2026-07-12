namespace ExcelCSIToolBox.RefBuilder.Parsing
{
    public sealed class ApiParameterDefinition
    {
        public string Name { get; set; }
        public string TypeName { get; set; }
        public bool IsByRef { get; set; }
        public bool IsOut { get; set; }
        public string Description { get; set; }
    }
}
