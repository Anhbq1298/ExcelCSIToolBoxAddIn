using ExcelCSIToolBoxAddIn.Common.Results;
using ExcelCSIToolBoxAddIn.Infrastructure.Etabs;
using ExcelCSIToolBoxAddIn.Infrastructure.Excel;

namespace ExcelCSIToolBoxAddIn.Core.Application
{
    public class SelectFramesFromExcelRangeByUniqueNameUseCase
    {
        private readonly IEtabsConnectionService _connectionService;
        private readonly IExcelSelectionService _excelSelectionService;

        public SelectFramesFromExcelRangeByUniqueNameUseCase(
            IEtabsConnectionService connectionService,
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
