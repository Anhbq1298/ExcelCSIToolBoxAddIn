using ExcelCSIToolBoxAddIn.Common.Results;
using ExcelCSIToolBoxAddIn.Infrastructure.Etabs;

namespace ExcelCSIToolBoxAddIn.Core.Application
{
    /// <summary>
    /// Application use case for loading CSI connection state for the toolbox shell.
    /// </summary>
    public class LoadCsiConnectionUseCase
    {
        private readonly ICsiConnectionService _connectionService;

        public LoadCsiConnectionUseCase(ICsiConnectionService connectionService)
        {
            _connectionService = connectionService;
        }

        public OperationResult<CsiConnectionInfo> Execute()
        {
            return _connectionService.TryAttachToRunningInstance();
        }
    }
}
