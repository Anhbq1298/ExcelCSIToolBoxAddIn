using System;
using System.Collections.Generic;

namespace ExcelCSIToolBox.AI.Mcp.Tools
{
    /// <summary>
    /// Concrete in-memory registry for read-only MCP tools.
    /// Refuses to register any tool where IsReadOnly = false.
    /// </summary>
    public class McpToolRegistry : IMcpToolRegistry
    {
        private readonly Dictionary<string, IMcpTool> _tools =
            new Dictionary<string, IMcpTool>(StringComparer.OrdinalIgnoreCase);

        /// <summary>
        /// Register a tool. Throws InvalidOperationException if the tool is not read-only.
        /// </summary>
        public void Register(IMcpTool tool)
        {
            if (tool == null)
            {
                throw new ArgumentNullException(nameof(tool));
            }

            // Safety rule: only read-only tools may be registered.
            if (!tool.IsReadOnly)
            {
                throw new InvalidOperationException(
                    $"Tool '{tool.Name}' is not read-only and cannot be registered in the LocalMcpServer.");
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
