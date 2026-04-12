using ExcelCSIToolBoxAddIn.Common.Results;
using ExcelCSIToolBoxAddIn.Infrastructure.Etabs;
using ExcelCSIToolBoxAddIn.Infrastructure.Excel;

namespace ExcelCSIToolBoxAddIn.Core.Application
{
    public class GetSelectedEtabsPointsUseCase
    {
        private readonly IEtabsConnectionService _connectionService;
        private readonly IExcelOutputService _excelOutputService;

        public GetSelectedEtabsPointsUseCase(
            IEtabsConnectionService connectionService,
            IExcelOutputService excelOutputService)
        {
            _connectionService = connectionService;
            _excelOutputService = excelOutputService;
        }

        public OperationResult Execute()
        {
            var pointsResult = _connectionService.GetSelectedPointsFromActiveModel();
            if (!pointsResult.IsSuccess)
            {
                return OperationResult.Failure(pointsResult.Message);
            }

            return _excelOutputService.WritePointsToActiveCell(pointsResult.Data);
        }
    }
}
