using System.Collections.Generic;
using ExcelCSIToolBox.Core.Common.Results;
using ExcelCSIToolBox.Core.Abstractions.CSI;

namespace ExcelCSIToolBox.Core.Application
{
    public class GetLoadPatternsUseCase
    {
        private readonly ICSISapModelConnectionService _connectionService;

        public GetLoadPatternsUseCase(ICSISapModelConnectionService connectionService)
        {
            _connectionService = connectionService;
        }

        public OperationResult<IReadOnlyList<ExcelCSIToolBox.Data.DTOs.CSI.CSISapModelLoadPatternDTO>> Execute()
        {
            return _connectionService.GetLoadPatterns();
        }
    }
}

