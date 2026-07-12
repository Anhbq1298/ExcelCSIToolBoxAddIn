using System.Collections.Generic;

namespace ExcelCSIToolBox.Core.Contracts.CSI
{
    public class LoadCombinationApplyResultDto
    {
        public int ProcessedCount { get; set; }
        public int SuccessCount { get; set; }
        public int FailedCount { get; set; }
        public List<LoadCombinationApplyFailureDto> Failures { get; set; } = new List<LoadCombinationApplyFailureDto>();
    }
}
