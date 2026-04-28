using System.Collections.Generic;
namespace ExcelCSIToolBox.AI.Mcp.Contracts {
    public class ToolResult<TData>
    {
        public bool Success { get; set; }
        public string Message { get; set; } = string.Empty;
        public string ErrorCode { get; set; } = string.Empty;
        public List<string> Warnings { get; set; } = new List<string>();
        public TData Data { get; set; }
    }
}
