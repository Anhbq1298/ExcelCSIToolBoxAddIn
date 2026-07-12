namespace ExcelCSIToolBox.AI.Agent
{
    /// <summary>
    /// Parsed decision from the tool-routing LLM call.
    /// If ShouldCallTool is true, the orchestrator calls the named tool.
    /// </summary>
    public class AiAgentToolDecision
    {
        /// <summary>Whether the AI decided a tool should be called.</summary>
        public bool ShouldCallTool { get; set; }

        /// <summary>Tool name to call (e.g. "CSI.GetModelInfo"). Empty if ShouldCallTool = false.</summary>
        public string ToolName { get; set; }

        /// <summary>JSON arguments for the tool. Usually "{}" for read-only no-arg tools.</summary>
        public string ArgumentsJson { get; set; }

        /// <summary>LLM's explanation of why this decision was made.</summary>
        public string Reason { get; set; }

        /// <summary>Whether the request was understood as model-related but is missing safe dispatch details.</summary>
        public bool ClarificationRequired { get; set; }

        /// <summary>User-facing clarification text. When set, no tool should be called.</summary>
        public string ClarificationMessage { get; set; }

        /// <summary>Candidate domain used when no executable MCP tool was selected.</summary>
        public string CandidateDomain { get; set; }

        /// <summary>Candidate action used when no executable MCP tool was selected.</summary>
        public string CandidateAction { get; set; }

        /// <summary>Candidate target object used when no executable MCP tool was selected.</summary>
        public string TargetObject { get; set; }

        /// <summary>User-facing diagnostic for a missing schema or tool route.</summary>
        public string MissingSchemaMessage { get; set; }
    }
}
