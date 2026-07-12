using System.Collections.Generic;

namespace ExcelCSIToolBox.Application.ToolCatalog.Contracts
{
    public sealed class ToolRequest
    {
        public string RawText { get; set; }
        public string Action { get; set; }
        public string TargetObject { get; set; }
        public Dictionary<string, string> Parameters { get; set; }
    }
}
