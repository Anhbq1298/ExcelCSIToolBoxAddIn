using System.Collections.Generic;
using ExcelCSIToolBox.Core.Common.Results;
using ExcelCSIToolBox.Core.Abstractions.CSI;

namespace ExcelCSIToolBox.Application.Features.Loadings
{
    public class GetLoadCombinationDetailsUseCase
    {
        private readonly ICSISapModelConnectionService _connectionService;

        public GetLoadCombinationDetailsUseCase(ICSISapModelConnectionService connectionService)
        {
            _connectionService = connectionService;
        }

        public OperationResult<IReadOnlyList<ExcelCSIToolBox.Core.Contracts.CSI.LoadCombinationItemDTO>> Execute(string combinationName)
        {
            var result = _connectionService.GetLoadCombinationDetails(combinationName);
            return result;
        }
    }
}

