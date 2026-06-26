using System;
using System.Collections.Generic;
using ExcelCSIToolBox.Core.Abstractions.CSI;
using ExcelCSIToolBox.Core.Common.Results;
using ExcelCSIToolBox.Data.DTOs.CSI;

namespace ExcelCSIToolBox.Application.UseCases
{
    public class GetStoryForcesUseCase
    {
        private readonly ICSISapModelConnectionService _csiConnectionService;

        public GetStoryForcesUseCase(ICSISapModelConnectionService csiConnectionService)
        {
            _csiConnectionService = csiConnectionService ?? throw new ArgumentNullException(nameof(csiConnectionService));
        }

        public OperationResult<IReadOnlyList<CSISapModelStoryForceRowDTO>> Execute(IReadOnlyList<CSISapModelOutputCaseDTO> selectedOutputCases)
        {
            return _csiConnectionService.GetStoryForces(selectedOutputCases);
        }
    }
}
