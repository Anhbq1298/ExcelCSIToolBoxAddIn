using System.Collections.Generic;

namespace ExcelCSIToolBox.AI.Agent
{
    public sealed class AgentTaskItem
    {
        public string Id { get; set; }
        public string OriginalText { get; set; }
        public string NormalizedIntent { get; set; }
        public string TargetObjectType { get; set; }
        public string ActionType { get; set; }
        public Dictionary<string, string> Parameters { get; set; }
        public string Status { get; set; }
        public string ResultMessage { get; set; }
        public bool NeedsClarification { get; set; }
    }
}
