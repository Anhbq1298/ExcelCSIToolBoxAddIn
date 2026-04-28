using ExcelCSIToolBoxAddIn.Common.Results;
using ExcelCSIToolBoxAddIn.Data;
using ExcelCSIToolBoxAddIn.Data.DTOs;
using ExcelCSIToolBoxAddIn.Data.Models;
using ExcelCSIToolBoxAddIn.Data.Mappers;
using ExcelCSIToolBoxAddIn.Infrastructure.CSISapModel;
using ExcelCSIToolBoxAddIn.Infrastructure.Excel;

namespace ExcelCSIToolBoxAddIn.Core.Application
{
    public class GetSelectedCSISapModelPointsUseCase
    {
        private readonly ICSISapModelConnectionService _connectionService;
        private readonly IExcelOutputService _excelOutputService;

        public GetSelectedCSISapModelPointsUseCase(
            ICSISapModelConnectionService connectionService,
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

            var dataFrame = CSISapModelPointDataDataFrameMapper.Map(pointsResult.Data);
            return _excelOutputService.WriteDataFrameToActiveCell(dataFrame);
        }
    }
}
