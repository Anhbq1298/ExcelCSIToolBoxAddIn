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
        private readonly Dictionary<string, string> _aliases =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

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

        public void RegisterAlias(string alias, string toolName)
        {
            if (string.IsNullOrWhiteSpace(alias) || string.IsNullOrWhiteSpace(toolName))
            {
                return;
            }

            _aliases[alias] = toolName;
        }

        /// <summary>Returns the tool with the given name, or null if not found.</summary>
        public IMcpTool GetTool(string toolName)
        {
            if (string.IsNullOrWhiteSpace(toolName))
            {
                return null;
            }

            IMcpTool tool;
            if (_tools.TryGetValue(toolName, out tool))
            {
                return tool;
            }

            string canonicalToolName;
            return _aliases.TryGetValue(toolName, out canonicalToolName) &&
                   _tools.TryGetValue(canonicalToolName, out tool)
                ? tool
                : null;
        }

        /// <summary>Returns all registered tools.</summary>
        public IReadOnlyList<IMcpTool> GetAllTools()
        {
            return new List<IMcpTool>(_tools.Values);
        }
    }
}
