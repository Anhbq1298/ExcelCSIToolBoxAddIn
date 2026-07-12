using ExcelCSIToolBox.Core.Common.Results;
using ExcelCSIToolBox.Data;
using ExcelCSIToolBox.Data.DTOs.CSI;
using ExcelCSIToolBox.Data.Models;
using ExcelCSIToolBox.Application.Mappers;
using ExcelCSIToolBox.Core.Abstractions.CSI;
using ExcelCSIToolBox.Core.Abstractions.Excel;

namespace ExcelCSIToolBox.Application.UseCases
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

