using System.Collections.Generic;

namespace ExcelCSIToolBox.Core.GenerativeDesign
{
    public sealed class EvaluationResult
    {
        public string OptionId { get; set; }
        public bool IsValid { get; set; }
        public double Score { get; set; }
        public double EstimatedDriftRatio { get; set; }
        public double EstimatedWeight { get; set; }
        public List<string> Messages { get; set; }

        public EvaluationResult()
        {
            Messages = new List<string>();
        }
    }
}
