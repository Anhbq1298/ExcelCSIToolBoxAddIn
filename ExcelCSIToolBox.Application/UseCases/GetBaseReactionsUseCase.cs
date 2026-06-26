using System;
using System.Collections.Generic;
using ExcelCSIToolBox.Core.Abstractions.CSI;
using ExcelCSIToolBox.Core.Common.Results;
using ExcelCSIToolBox.Data.DTOs.CSI;

namespace ExcelCSIToolBox.Application.UseCases
{
    public class GetBaseReactionsUseCase
    {
        private readonly ICSISapModelConnectionService _csiConnectionService;

        public GetBaseReactionsUseCase(ICSISapModelConnectionService csiConnectionService)
        {
            _csiConnectionService = csiConnectionService ?? throw new ArgumentNullException(nameof(csiConnectionService));
        }

        public OperationResult<IReadOnlyList<CSISapModelBaseReactionRowDTO>> Execute(IReadOnlyList<CSISapModelOutputCaseDTO> selectedOutputCases)
        {
            return _csiConnectionService.GetBaseReactions(selectedOutputCases);
        }
    }
}
