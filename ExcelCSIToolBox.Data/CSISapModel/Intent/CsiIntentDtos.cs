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
}
