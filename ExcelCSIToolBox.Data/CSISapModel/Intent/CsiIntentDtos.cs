using System.Collections.Generic;

namespace ExcelCSIToolBox.Data.CSISapModel.Intent
{
    public sealed class CsiIntentPlanDto
    {
        public List<CsiIntentTaskDto> Tasks { get; set; }
    }

    public sealed class CsiIntentTaskDto
    {
        public string TaskType { get; set; }
        public string Operation { get; set; }
        public Dictionary<string, string> Arguments { get; set; }
        public List<string> DependsOn { get; set; }
    }

    public sealed class CsiRequestClassificationDto
    {
        public string Status { get; set; }
        public List<CsiRequestTaskClassificationDto> Tasks { get; set; }
    }

    public sealed class CsiRequestTaskClassificationDto
    {
        public string RawText { get; set; }
        public string Action { get; set; }
        public string TargetObject { get; set; }
        public Dictionary<string, string> Parameters { get; set; }
        public List<string> MissingParameters { get; set; }
        public string ToolName { get; set; }
        public string ClarificationMessage { get; set; }
    }
}
