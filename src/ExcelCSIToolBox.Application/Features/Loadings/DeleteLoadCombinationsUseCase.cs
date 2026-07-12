using System.Collections.Generic;
using ExcelCSIToolBox.Core.Common.Results;
using ExcelCSIToolBox.Core.Abstractions.CSI;

namespace ExcelCSIToolBox.Application.Features.Loadings
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

