using ExcelCSIToolBox.Core.Models.CSI;

namespace ExcelCSIToolBox.AI.Mcp.Contracts
{
    public sealed class McpToolDescriptor
    {
        public string Name { get; set; }
        public string Title { get; set; }
        public string Category { get; set; }
        public string SubCategory { get; set; }
        public string Description { get; set; }
        public bool IsReadOnly { get; set; }
        public CsiMethodRiskLevel RiskLevel { get; set; }
        public bool RequiresConfirmation { get; set; }
        public bool SupportsDryRun { get; set; }
    }
}
