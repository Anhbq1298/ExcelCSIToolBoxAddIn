using ExcelCSIToolBoxAddIn.Common.Results;
using ExcelCSIToolBoxAddIn.Infrastructure.Etabs;
using ExcelCSIToolBoxAddIn.Infrastructure.Excel;

namespace ExcelCSIToolBoxAddIn.Core.Application
{
    public class GetSelectedCsiFramesUseCase
    {
        private readonly ICsiConnectionService _connectionService;
        private readonly IExcelOutputService _excelOutputService;

        public GetSelectedCsiFramesUseCase(
            ICsiConnectionService connectionService,
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

            var dataFrame = CsiFrameDataDataFrameMapper.Map(framesResult.Data);
            return _excelOutputService.WriteDataFrameToActiveCell(dataFrame);
        }
    }
}
