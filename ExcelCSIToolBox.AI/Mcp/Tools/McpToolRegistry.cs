using System;
using System.Collections.Generic;

namespace ExcelCSIToolBox.AI.Mcp.Tools
{
    /// <summary>
    /// Concrete in-memory registry for local MCP tools.
    /// Write tools are allowed only when implemented as typed, guarded tools.
    /// </summary>
    public class McpToolRegistry : IMcpToolRegistry
    {
        private readonly Dictionary<string, IMcpTool> _tools =
            new Dictionary<string, IMcpTool>(StringComparer.OrdinalIgnoreCase);

        /// <summary>
        /// Register a typed tool.
        /// </summary>
        public void Register(IMcpTool tool)
        {
            if (tool == null)
            {
                throw new ArgumentNullException(nameof(tool));
            }

            _tools[tool.Name] = tool;
        }

        /// <summary>Returns the tool with the given name, or null if not found.</summary>
        public IMcpTool GetTool(string toolName)
        {
            if (string.IsNullOrWhiteSpace(toolName))
            {
                return null;
            }

            IMcpTool tool;
            return _tools.TryGetValue(toolName, out tool) ? tool : null;
        }

        /// <summary>Returns all registered tools.</summary>
        public IReadOnlyList<IMcpTool> GetAllTools()
        {
            return new List<IMcpTool>(_tools.Values);
        }
    }
}
