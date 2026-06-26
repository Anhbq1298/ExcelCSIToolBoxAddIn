using System;
using ExcelCSIToolBox.Core.Abstractions.CSI;
using ExcelCSIToolBox.Core.Common.Results;
using ExcelCSIToolBox.Data.DTOs.CSI;

namespace ExcelCSIToolBox.Application.UseCases
{
    public class GetMassSummaryByStoryUseCase
    {
        private readonly ICSISapModelConnectionService _csiConnectionService;

        public GetMassSummaryByStoryUseCase(ICSISapModelConnectionService csiConnectionService)
        {
            _csiConnectionService = csiConnectionService ?? throw new ArgumentNullException(nameof(csiConnectionService));
        }

        public OperationResult<CSISapModelDisplayTableDTO> Execute()
        {
            return _csiConnectionService.GetMassSummaryByStory();
        }
    }
}
