using System;
using System.Collections.Generic;
using ExcelCSIToolBox.Core.Abstractions.CSI;
using ExcelCSIToolBox.Core.Common.Results;
using ExcelCSIToolBox.Core.Contracts.CSI;

namespace ExcelCSIToolBox.Application.UseCases
{
    public class GetStoryDriftsUseCase
    {
        private readonly ICSISapModelConnectionService _csiConnectionService;

        public GetStoryDriftsUseCase(ICSISapModelConnectionService csiConnectionService)
        {
            _csiConnectionService = csiConnectionService ?? throw new ArgumentNullException(nameof(csiConnectionService));
        }

        public OperationResult<CSISapModelDisplayTableDTO> Execute(IReadOnlyList<CSISapModelOutputCaseDTO> selectedOutputCases)
        {
            return _csiConnectionService.GetStoryDrifts(selectedOutputCases);
        }
    }
}
