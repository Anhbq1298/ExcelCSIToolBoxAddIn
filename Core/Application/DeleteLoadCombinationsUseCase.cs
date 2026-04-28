using System.Collections.Generic;
using ExcelCSIToolBoxAddIn.Common.Results;
using ExcelCSIToolBoxAddIn.Infrastructure.CSISapModel;

namespace ExcelCSIToolBoxAddIn.Core.Application
{
    public class DeleteLoadCombinationsUseCase
    {
        private readonly ICSISapModelConnectionService _connectionService;

        public DeleteLoadCombinationsUseCase(ICSISapModelConnectionService connectionService)
        {
            _connectionService = connectionService;
        }

        public OperationResult Execute(IReadOnlyList<string> combinationNames)
        {
            return _connectionService.DeleteLoadCombinations(combinationNames);
        }
    }
}
