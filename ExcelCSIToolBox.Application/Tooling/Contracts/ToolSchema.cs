using System.Collections.Generic;

namespace ExcelCSIToolBox.Application.Tooling.Contracts
{
    public sealed class ToolSchema
    {
        public string Domain { get; set; }
        public string ToolName { get; set; }
        public List<string> Aliases { get; set; }
        public string Action { get; set; }
        public string TargetObject { get; set; }
        public List<ToolParameterSchema> RequiredParameters { get; set; }
        public List<ToolParameterSchema> OptionalParameters { get; set; }
        public List<string> IntentHints { get; set; }
        public string ClarificationMessage { get; set; }
        public bool IsModelMutating { get; set; }
    }
}
