using System.Collections.Generic;
using ExcelCSIToolBoxAddIn.Common.Results;
using ExcelCSIToolBoxAddIn.Infrastructure.CSISapModel;

namespace ExcelCSIToolBoxAddIn.Core.Application
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
