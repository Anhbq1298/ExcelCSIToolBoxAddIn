using ExcelCSIToolBox.Core.Common.Results;
using ExcelCSIToolBox.Core.Abstractions.CSI;

namespace ExcelCSIToolBox.Core.Application
{
    public class CloseCurrentInstanceUseCase
    {
        private readonly ICSISapModelConnectionService _connectionService;

        public CloseCurrentInstanceUseCase(ICSISapModelConnectionService connectionService)
        {
            _connectionService = connectionService;
        }

        public OperationResult Execute()
        {
            return _connectionService.CloseCurrentInstance();
        }
    }
}

