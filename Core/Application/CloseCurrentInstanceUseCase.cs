using ExcelCSIToolBoxAddIn.Common.Results;
using ExcelCSIToolBoxAddIn.Infrastructure.Csi;

namespace ExcelCSIToolBoxAddIn.Core.Application
{
    public class CloseCurrentInstanceUseCase
    {
        private readonly ICsiConnectionService _connectionService;

        public CloseCurrentInstanceUseCase(ICsiConnectionService connectionService)
        {
            _connectionService = connectionService;
        }

        public OperationResult Execute()
        {
            return _connectionService.CloseCurrentInstance();
        }
    }
}
