using System.Collections.Generic;

namespace ExcelCSIToolBox.AI.Mcp.Tools
{
    /// <summary>
    /// Registry that holds all registered MCP tools.
    /// Only read-only tools may be registered.
    /// </summary>
    public interface IMcpToolRegistry
    {
        /// <summary>Register a tool. Throws if the tool is not read-only.</summary>
        void Register(IMcpTool tool);

        /// <summary>Retrieve a tool by its unique name. Returns null if not found.</summary>
        IMcpTool GetTool(string toolName);

        /// <summary>Returns all currently registered tools.</summary>
        IReadOnlyList<IMcpTool> GetAllTools();
    }
}
