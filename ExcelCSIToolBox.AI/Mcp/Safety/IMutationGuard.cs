using System.Threading;
using System.Threading.Tasks;

namespace ExcelCSIToolBox.AI.Mcp.Safety
{
    /// <summary>
    /// Confirms whether a proposed write-path MCP operation may execute.
    /// </summary>
    public interface IMutationGuard
    {
        /// <summary>
        /// Returns true if the user confirmed the proposed write operation.
        /// </summary>
        Task<bool> ConfirmAsync(string toolName, string summary, CancellationToken ct);
    }
}
