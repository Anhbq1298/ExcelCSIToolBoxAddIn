using System.Collections.Generic;
using ExcelCSIToolBox.Core.Common.Results;
using ExcelCSIToolBox.Core.Abstractions.CSI;

namespace ExcelCSIToolBox.Application.UseCases
{
    public class GetLoadCombinationsUseCase
    {
        private readonly ICSISapModelConnectionService _connectionService;

        public GetLoadCombinationsUseCase(ICSISapModelConnectionService connectionService)
        {
            _connectionService = connectionService;
        }

        public OperationResult<IReadOnlyList<ExcelCSIToolBox.Data.DTOs.CSI.CSISapModelLoadCombinationDTO>> Execute()
        {
            var result = _connectionService.GetLoadCombinations();
            return result;
        }
    }
}

