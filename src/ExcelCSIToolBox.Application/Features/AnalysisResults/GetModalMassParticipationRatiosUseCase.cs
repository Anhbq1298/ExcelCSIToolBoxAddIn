using System;
using System.Collections.Generic;
using ExcelCSIToolBox.Core.Abstractions.CSI;
using ExcelCSIToolBox.Core.Common.Results;
using ExcelCSIToolBox.Core.Contracts.CSI;

namespace ExcelCSIToolBox.Application.Features.AnalysisResults
{
    public class GetModalMassParticipationRatiosUseCase
    {
        private readonly ICSISapModelConnectionService _csiConnectionService;

        public GetModalMassParticipationRatiosUseCase(ICSISapModelConnectionService csiConnectionService)
        {
            _csiConnectionService = csiConnectionService ?? throw new ArgumentNullException(nameof(csiConnectionService));
        }

        public OperationResult<IReadOnlyList<CSISapModelModalMassParticipationRowDTO>> Execute(IReadOnlyList<CSISapModelOutputCaseDTO> selectedLoadCases)
        {
            return _csiConnectionService.GetModalMassParticipationRatios(selectedLoadCases);
        }
    }
}
