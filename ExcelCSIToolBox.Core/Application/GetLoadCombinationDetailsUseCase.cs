using System.Collections.Generic;
using ExcelCSIToolBox.Core.Common.Results;
using ExcelCSIToolBox.Core.Abstractions.CSI;

namespace ExcelCSIToolBox.Core.Application
{
    public class GetLoadCombinationDetailsUseCase
    {
        private readonly ICSISapModelConnectionService _connectionService;

        public GetLoadCombinationDetailsUseCase(ICSISapModelConnectionService connectionService)
        {
            _connectionService = connectionService;
        }

        public OperationResult<IReadOnlyList<ExcelCSIToolBox.Data.DTOs.CSI.LoadCombinationItemDTO>> Execute(string combinationName)
        {
            var result = _connectionService.GetLoadCombinationDetails(combinationName);
            return result;
        }
    }
}

