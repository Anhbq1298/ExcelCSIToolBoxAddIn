using ExcelCSIToolBoxAddIn.Common.Results;
using ExcelCSIToolBoxAddIn.Infrastructure.Csi;
using ExcelCSIToolBoxAddIn.Infrastructure.Excel;

namespace ExcelCSIToolBoxAddIn.Core.Application
{
    public class SelectFramesFromExcelRangeByUniqueNameUseCase
    {
        private readonly ICsiConnectionService _connectionService;
        private readonly IExcelSelectionService _excelSelectionService;

        public SelectFramesFromExcelRangeByUniqueNameUseCase(
            ICsiConnectionService connectionService,
            IExcelSelectionService excelSelectionService)
        {
            _connectionService = connectionService;
            _excelSelectionService = excelSelectionService;
        }

        public OperationResult Execute()
        {
            var valuesResult = _excelSelectionService.ReadSingleColumnTextValues();
            if (!valuesResult.IsSuccess)
            {
                return OperationResult.Failure(valuesResult.Message);
            }

            return _connectionService.SelectFramesByUniqueNames(valuesResult.Data);
        }
    }
}
