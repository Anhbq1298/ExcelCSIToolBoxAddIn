using System.Collections.Generic;
using ExcelCSIToolBoxAddIn.Common.Results;
using ExcelCSIToolBoxAddIn.Infrastructure.CSISapModel;

namespace ExcelCSIToolBoxAddIn.Core.Application
{
    public class GetLoadCombinationsUseCase
    {
        private readonly ICSISapModelConnectionService _connectionService;

        public GetLoadCombinationsUseCase(ICSISapModelConnectionService connectionService)
        {
            _connectionService = connectionService;
        }

        public OperationResult<IReadOnlyList<string>> Execute()
        {
            return _connectionService.GetLoadCombinations();
        }
    }
}
