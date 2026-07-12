using ExcelCSIToolBox.Core.Models.CSI;

namespace ExcelCSIToolBox.AI.Mcp.Tools
{
    public interface IMcpToolMetadata
    {
        string Title { get; }
        string Category { get; }
        string SubCategory { get; }
        CsiMethodRiskLevel RiskLevel { get; }
        bool RequiresConfirmation { get; }
        bool SupportsDryRun { get; }
    }
}
