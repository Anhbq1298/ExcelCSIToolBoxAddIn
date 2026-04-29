using System.Collections.Generic;

namespace ExcelCSIToolBox.Data.CSISapModel.Workflow
{
    public sealed class CsiWorkflowRequestDto
    {
        public string UserInput { get; set; }
        public List<CsiTaskDto> PlannedTasks { get; set; }
    }

    public sealed class CsiTaskDto
    {
        public string TaskId { get; set; }
        public string TaskType { get; set; }
        public string Operation { get; set; }
        public Dictionary<string, string> Arguments { get; set; }
        public List<string> DependsOn { get; set; }
    }

    public sealed class CsiTaskResultDto
    {
        public string TaskId { get; set; }
        public string TaskType { get; set; }
        public string Operation { get; set; }
        public bool Success { get; set; }
        public bool Skipped { get; set; }
        public string ObjectName { get; set; }
        public string Message { get; set; }
        public string FailureReason { get; set; }
    }

    public sealed class CsiWorkflowResultDto
    {
        public int TotalTasksDetected { get; set; }
        public int Succeeded { get; set; }
        public int Failed { get; set; }
        public int Skipped { get; set; }
        public List<CsiTaskResultDto> Results { get; set; }
    }
}
