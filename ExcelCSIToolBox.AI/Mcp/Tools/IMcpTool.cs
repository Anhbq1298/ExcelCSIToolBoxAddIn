using System.Threading;
using System.Threading.Tasks;
using ExcelCSIToolBox.AI.Mcp.Contracts;

namespace ExcelCSIToolBox.AI.Mcp.Tools
{
    /// <summary>
    /// Contract for a single local MCP-style tool the AI agent can call.
    /// All tools used by the AI agent must implement this interface.
    /// </summary>
    public interface IMcpTool
    {
        /// <summary>Unique tool identifier used by the AI agent, e.g. "CSI.GetModelInfo".</summary>
        string Name { get; }

        /// <summary>Human-readable description shown to the tool router LLM.</summary>
        string Description { get; }

        /// <summary>True for read-only query tools; false for guarded write tools.</summary>
        bool IsReadOnly { get; }

        /// <summary>Execute the tool and return a structured response.</summary>
        Task<ToolCallResponse> ExecuteAsync(string argumentsJson, CancellationToken cancellationToken);
    }
}
