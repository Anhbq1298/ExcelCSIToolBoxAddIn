namespace ExcelCSIToolBox.AI.Agent
{
    public sealed class AiAgentResponseBuilder
    {
        public AiAgentResponse Clarification(string message, string reason)
        {
            return new AiAgentResponse
            {
                AssistantText = message,
                ToolWasCalled = false,
                RoutingReason = reason
            };
        }
    }
}
