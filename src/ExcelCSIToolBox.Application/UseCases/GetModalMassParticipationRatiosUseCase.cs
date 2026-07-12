using System;
using System.Collections.Generic;
using ExcelCSIToolBox.Core.Abstractions.CSI;
using ExcelCSIToolBox.Core.Common.Results;
using ExcelCSIToolBox.Data.DTOs.CSI;

namespace ExcelCSIToolBox.Application.UseCases
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
