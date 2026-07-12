using System;
using ExcelCSIToolBox.Application.Interfaces.Etabs;
using ExcelCSIToolBox.Core.Abstractions.CSI;
using ExcelCSIToolBox.Core.Common.Results;
using ExcelCSIToolBox.Core.Contracts.CSI;

namespace ExcelCSIToolBox.Infrastructure.Services.Etabs
{
    public class EtabsConnectionService : IEtabsConnectionService
    {
        private readonly ICSISapModelConnectionService _connectionService;

        public EtabsConnectionService(ICSISapModelConnectionService connectionService)
        {
            _connectionService = connectionService ?? throw new ArgumentNullException(nameof(connectionService));
        }

        public object SapModel
        {
            get
            {
                OperationResult<CSISapModelConnectionInfoDTO> result = _connectionService.GetCurrentConnection();
                return result.IsSuccess && result.Data != null ? result.Data.SapModel : null;
            }
        }

        public OperationResult<CSISapModelPresentUnitSystemDTO> GetPresentUnitSystem()
        {
            return _connectionService.GetPresentUnitSystem();
        }

        public OperationResult SetPresentUnitSystem(CSISapModelPresentUnitSystemDTO unitSystem)
        {
            return _connectionService.SetPresentUnitSystem(unitSystem);
        }
    }
}
