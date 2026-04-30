using ExcelCSIToolBox.AI.Mcp.Contracts;

namespace ExcelCSIToolBox.AI.Agent
{
    /// <summary>
    /// The complete response produced by the AI agent orchestrator for one user message.
    /// </summary>
    public class AiAgentResponse
    {
        /// <summary>Final assistant text shown to the user.</summary>
        public string AssistantText { get; set; }

        /// <summary>Whether a tool was called during this response.</summary>
        public bool ToolWasCalled { get; set; }

        /// <summary>Name of the tool that was called (null if none).</summary>
        public string ToolName { get; set; }

        /// <summary>Arguments JSON sent to the tool (null if none).</summary>
        public string ToolArgumentsJson { get; set; }

        /// <summary>Structured tool response (null if no tool was called).</summary>
        public ToolCallResponse ToolResponse { get; set; }

        /// <summary>LLM's routing reason (debug/trace text).</summary>
        public string RoutingReason { get; set; }
    }
}
