using ExcelCSIToolBoxAddIn.Common.Results;
using ExcelCSIToolBoxAddIn.Infrastructure.Etabs;

namespace ExcelCSIToolBoxAddIn.Core.Application
{
    public class CloseCurrentEtabsInstanceUseCase
    {
        private readonly IEtabsConnectionService _connectionService;

        public CloseCurrentEtabsInstanceUseCase(IEtabsConnectionService connectionService)
        {
            _connectionService = connectionService;
        }

        public OperationResult Execute()
        {
            return _connectionService.CloseCurrentEtabsInstance();
        }
    }
}
