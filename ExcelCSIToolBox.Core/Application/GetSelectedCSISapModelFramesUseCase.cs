using ExcelCSIToolBox.Core.Common.Results;
using ExcelCSIToolBox.Data;
using ExcelCSIToolBox.Data.DTOs.CSI;
using ExcelCSIToolBox.Data.Models;
using ExcelCSIToolBox.Data.Mappers;
using ExcelCSIToolBox.Core.Abstractions.CSI;
using ExcelCSIToolBox.Core.Abstractions.Excel;

namespace ExcelCSIToolBox.Core.Application
{
    public class GetSelectedCSISapModelFramesUseCase
    {
        private readonly ICSISapModelConnectionService _connectionService;
        private readonly IExcelOutputService _excelOutputService;

        public GetSelectedCSISapModelFramesUseCase(
            ICSISapModelConnectionService connectionService,
            IExcelOutputService excelOutputService)
        {
            _connectionService = connectionService;
            _excelOutputService = excelOutputService;
        }

        public OperationResult Execute()
        {
            var framesResult = _connectionService.GetSelectedFramesFromActiveModel();
            if (!framesResult.IsSuccess)
            {
                return OperationResult.Failure(framesResult.Message);
            }

            var dataFrame = CSISapModelFrameDataDataFrameMapper.Map(framesResult.Data);
            return _excelOutputService.WriteDataFrameToActiveCell(dataFrame);
        }
    }
}

