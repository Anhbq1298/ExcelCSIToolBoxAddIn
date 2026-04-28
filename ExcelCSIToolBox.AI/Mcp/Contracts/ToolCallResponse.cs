namespace ExcelCSIToolBox.AI.Mcp.Contracts
{
    /// <summary>
    /// Structured response returned after a local MCP tool has been executed.
    /// </summary>
    public class ToolCallResponse
    {
        /// <summary>The tool that was called.</summary>
        public string ToolName { get; set; }

        /// <summary>Whether the tool completed successfully.</summary>
        public bool Success { get; set; }

        /// <summary>Human-readable success or error message.</summary>
        public string Message { get; set; }

        /// <summary>JSON-encoded result payload (may be null on failure).</summary>
        public string ResultJson { get; set; }
    }
}
