using System;
using System.Collections.Generic;
using ExcelCSIToolBox.Core.Abstractions.CSI;
using ExcelCSIToolBox.Core.Common.Results;
using ExcelCSIToolBox.Data.DTOs.CSI;

namespace ExcelCSIToolBox.Application.UseCases
{
    public class GetStoryDisplacementsUseCase
    {
        private readonly ICSISapModelConnectionService _csiConnectionService;

        public GetStoryDisplacementsUseCase(ICSISapModelConnectionService csiConnectionService)
        {
            _csiConnectionService = csiConnectionService ?? throw new ArgumentNullException(nameof(csiConnectionService));
        }

        public OperationResult<IReadOnlyList<CSISapModelStoryDisplacementRowDTO>> Execute(IReadOnlyList<CSISapModelOutputCaseDTO> selectedOutputCases)
        {
            return _csiConnectionService.GetStoryDisplacements(selectedOutputCases);
        }
    }
}
