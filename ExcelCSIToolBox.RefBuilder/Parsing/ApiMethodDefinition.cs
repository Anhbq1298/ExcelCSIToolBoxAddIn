using System.Collections.Generic;

namespace ExcelCSIToolBox.RefBuilder.Parsing
{
    public sealed class ApiMethodDefinition
    {
        public string ProductName { get; set; }
        public string ObjectName { get; set; }
        public string InterfaceName { get; set; }
        public string MethodName { get; set; }
        public string ReturnType { get; set; }
        public string FullSignature { get; set; }
        public string Category { get; set; }
        public string SafetyFlag { get; set; }
        public string SourceDocumentationFile { get; set; }
        public List<ApiParameterDefinition> Parameters { get; set; } = new List<ApiParameterDefinition>();
    }
}
