using ExcelCSIToolBoxAddIn.Common.Results;
using ExcelCSIToolBoxAddIn.Infrastructure.CSISapModel;
using ExcelCSIToolBoxAddIn.Infrastructure.Excel;

namespace ExcelCSIToolBoxAddIn.Core.Application
{
    public class SelectFramesFromExcelRangeByUniqueNameUseCase
    {
        private readonly ICSISapModelConnectionService _connectionService;
        private readonly IExcelSelectionService _excelSelectionService;

        public SelectFramesFromExcelRangeByUniqueNameUseCase(
            ICSISapModelConnectionService connectionService,
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
