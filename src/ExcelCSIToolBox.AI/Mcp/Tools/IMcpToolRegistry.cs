using System.Collections.Generic;

namespace ExcelCSIToolBox.AI.Mcp.Tools
{
    /// <summary>
    /// Registry that holds all registered typed MCP tools.
    /// </summary>
    public interface IMcpToolRegistry
    {
        /// <summary>Register a typed tool.</summary>
        void Register(IMcpTool tool);

        /// <summary>Register an alternate name that resolves to an existing typed tool.</summary>
        void RegisterAlias(string alias, string toolName);

        /// <summary>Retrieve a tool by its unique name. Returns null if not found.</summary>
        IMcpTool GetTool(string toolName);

        /// <summary>Returns all currently registered tools.</summary>
        IReadOnlyList<IMcpTool> GetAllTools();
    }
}
