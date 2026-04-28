using System.Collections.Generic;
using ExcelCSIToolBox.Core.Common.Results;
using ExcelCSIToolBox.Core.Abstractions.CSI;

namespace ExcelCSIToolBox.Core.Application
{
    public class DeleteLoadPatternsUseCase
    {
        private readonly ICSISapModelConnectionService _connectionService;

        public DeleteLoadPatternsUseCase(ICSISapModelConnectionService connectionService)
        {
            _connectionService = connectionService;
        }

        public OperationResult Execute(IReadOnlyList<string> patternNames)
        {
            return _connectionService.DeleteLoadPatterns(patternNames);
        }
    }
}

