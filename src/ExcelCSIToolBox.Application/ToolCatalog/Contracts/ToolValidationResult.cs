using System.Collections.Generic;

namespace ExcelCSIToolBox.Application.ToolCatalog.Contracts
{
    public class ToolValidationResult
    {
        public bool IsValid { get; set; }
        public ToolSchema Schema { get; set; }
        public string ToolName { get; set; }
        public List<string> MissingParameters { get; set; }
        public string ClarificationMessage { get; set; }
    }
}
