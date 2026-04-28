namespace ExcelCSIToolBox.AI.Mcp.Contracts
{
    /// <summary>
    /// Represents a request for an AI agent to call a local MCP-style tool.
    /// </summary>
    public class ToolCallRequest
    {
        /// <summary>The unique tool name, e.g. "CSI.GetModelInfo".</summary>
        public string ToolName { get; set; }

        /// <summary>JSON-encoded arguments for the tool (can be "{}" for no-arg tools).</summary>
        public string ArgumentsJson { get; set; }
    }
}
