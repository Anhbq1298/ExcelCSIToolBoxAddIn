using ExcelCSIToolBox.Core.Common.Results;
using ExcelCSIToolBox.Core.Abstractions.CSI;
using ExcelCSIToolBox.Core.Abstractions.Excel;

namespace ExcelCSIToolBox.Application.UseCases
{
    public class SelectPointsFromExcelRangeByUniqueNameUseCase
    {
        private readonly ICSISapModelConnectionService _connectionService;
        private readonly IExcelSelectionService _excelSelectionService;

        public SelectPointsFromExcelRangeByUniqueNameUseCase(
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

            return _connectionService.SelectPointsByUniqueNames(valuesResult.Data);
        }
    }
}

