using System;
using System.Collections.Generic;
using ExcelCSIToolBox.Core.Abstractions.CSI;
using ExcelCSIToolBox.Core.Common.Results;
using ExcelCSIToolBox.Core.Contracts.CSI;

namespace ExcelCSIToolBox.Application.UseCases
{
    public class GetStoryMaxOverAverageDisplacementsUseCase
    {
        private readonly ICSISapModelConnectionService _csiConnectionService;

        public GetStoryMaxOverAverageDisplacementsUseCase(ICSISapModelConnectionService csiConnectionService)
        {
            _csiConnectionService = csiConnectionService ?? throw new ArgumentNullException(nameof(csiConnectionService));
        }

        public OperationResult<CSISapModelDisplayTableDTO> Execute(IReadOnlyList<CSISapModelOutputCaseDTO> selectedOutputCases)
        {
            return _csiConnectionService.GetStoryMaxOverAverageDisplacements(selectedOutputCases);
        }
    }
}
