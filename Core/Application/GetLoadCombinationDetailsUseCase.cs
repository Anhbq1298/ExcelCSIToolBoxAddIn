using System.Collections.Generic;
using ExcelCSIToolBoxAddIn.Common.Results;
using ExcelCSIToolBoxAddIn.Infrastructure.CSISapModel;

namespace ExcelCSIToolBoxAddIn.Core.Application
{
    public class GetLoadCombinationDetailsUseCase
    {
        private readonly ICSISapModelConnectionService _connectionService;

        public GetLoadCombinationDetailsUseCase(ICSISapModelConnectionService connectionService)
        {
            _connectionService = connectionService;
        }

        public OperationResult<IReadOnlyList<ExcelCSIToolBoxAddIn.Data.DTOs.LoadCombinationItemDTO>> Execute(string combinationName)
        {
            var result = _connectionService.GetLoadCombinationDetails(combinationName);
            return result;
        }
    }
}
