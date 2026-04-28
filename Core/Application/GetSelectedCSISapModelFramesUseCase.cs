using ExcelCSIToolBoxAddIn.Common.Results;
using ExcelCSIToolBoxAddIn.Infrastructure.CSISapModel;
using ExcelCSIToolBoxAddIn.Infrastructure.Excel;

namespace ExcelCSIToolBoxAddIn.Core.Application
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
