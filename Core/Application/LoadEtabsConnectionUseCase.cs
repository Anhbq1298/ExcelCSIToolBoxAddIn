using ExcelCSIToolBoxAddIn.Common.Results;
using ExcelCSIToolBoxAddIn.Infrastructure.Etabs;

namespace ExcelCSIToolBoxAddIn.Core.Application
{
    /// <summary>
    /// Application use case for loading ETABS connection state for the toolbox shell.
    /// </summary>
    public class LoadEtabsConnectionUseCase
    {
        private readonly IEtabsConnectionService _connectionService;

        public LoadEtabsConnectionUseCase(IEtabsConnectionService connectionService)
        {
            _connectionService = connectionService;
        }

        public OperationResult<EtabsConnectionInfo> Execute()
        {
            return _connectionService.TryAttachToRunningInstance();
        }
    }
}
