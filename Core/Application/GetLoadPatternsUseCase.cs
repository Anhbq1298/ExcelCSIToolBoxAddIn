using System.Collections.Generic;
using ExcelCSIToolBoxAddIn.Common.Results;
using ExcelCSIToolBoxAddIn.Infrastructure.CSISapModel;

namespace ExcelCSIToolBoxAddIn.Core.Application
{
    public class GetLoadPatternsUseCase
    {
        private readonly ICSISapModelConnectionService _connectionService;

        public GetLoadPatternsUseCase(ICSISapModelConnectionService connectionService)
        {
            _connectionService = connectionService;
        }

        public OperationResult<IReadOnlyList<string>> Execute()
        {
            return _connectionService.GetLoadPatterns();
        }
    }
}
