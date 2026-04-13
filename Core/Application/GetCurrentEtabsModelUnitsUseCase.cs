using ExcelCSIToolBoxAddIn.Common.Results;
using ExcelCSIToolBoxAddIn.Infrastructure.Etabs;

namespace ExcelCSIToolBoxAddIn.Core.Application
{
    /// <summary>
    /// Application use case for reading the current present model units from ETABS.
    /// </summary>
    public class GetCurrentEtabsModelUnitsUseCase
    {
        private readonly IEtabsConnectionService _connectionService;

        public GetCurrentEtabsModelUnitsUseCase(IEtabsConnectionService connectionService)
        {
            _connectionService = connectionService;
        }

        public OperationResult<string> Execute()
        {
            return _connectionService.GetCurrentModelUnitsDisplayText();
        }
    }
}
