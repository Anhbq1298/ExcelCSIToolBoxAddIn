using ExcelCSIToolBoxAddIn.Common.Results;
using ExcelCSIToolBoxAddIn.Infrastructure.Etabs;

namespace ExcelCSIToolBoxAddIn.Core.Application
{
    /// <summary>
    /// Application use case for reading attached ETABS model display information.
    /// </summary>
    public class GetAttachedEtabsModelInfoUseCase
    {
        private readonly IEtabsConnectionService _connectionService;

        public GetAttachedEtabsModelInfoUseCase(IEtabsConnectionService connectionService)
        {
            _connectionService = connectionService;
        }

        public OperationResult<EtabsAttachedModelInfo> Execute()
        {
            return _connectionService.GetAttachedModelInfo();
        }
    }
}
