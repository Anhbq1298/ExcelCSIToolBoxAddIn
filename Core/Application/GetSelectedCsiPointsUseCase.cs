using ExcelCSIToolBoxAddIn.Common.Results;
using ExcelCSIToolBoxAddIn.Infrastructure.Csi;
using ExcelCSIToolBoxAddIn.Infrastructure.Excel;

namespace ExcelCSIToolBoxAddIn.Core.Application
{
    public class GetSelectedCsiPointsUseCase
    {
        private readonly ICsiConnectionService _connectionService;
        private readonly IExcelOutputService _excelOutputService;

        public GetSelectedCsiPointsUseCase(
            ICsiConnectionService connectionService,
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

            var dataFrame = CsiPointDataDataFrameMapper.Map(pointsResult.Data);
            return _excelOutputService.WriteDataFrameToActiveCell(dataFrame);
        }
    }
}
