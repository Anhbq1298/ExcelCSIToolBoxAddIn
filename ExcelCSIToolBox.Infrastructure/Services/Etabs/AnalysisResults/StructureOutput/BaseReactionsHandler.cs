using System.Threading.Tasks;
using ExcelCSIToolBox.Application.Interfaces.Etabs;
using ExcelCSIToolBox.Application.Interfaces.Etabs.AnalysisResults;
using ExcelCSIToolBox.Application.Interfaces.Excel;
using ExcelCSIToolBox.Core.Models.AnalysisResults;

namespace ExcelCSIToolBox.Infrastructure.Services.Etabs.AnalysisResults.StructureOutput
{
    public class BaseReactionsHandler : IEtabsAnalysisResultHandler
    {
        private readonly IEtabsDatabaseTableService _tableService;
        private readonly IExcelExportService _excelExportService;
        private readonly IEtabsUnitService _unitService;

        public BaseReactionsHandler(
            IEtabsDatabaseTableService tableService,
            IExcelExportService excelExportService,
            IEtabsUnitService unitService)
        {
            _tableService = tableService;
            _excelExportService = excelExportService;
            _unitService = unitService;
        }

        public bool CanHandle(string key)
        {
            return key == "BASE_REACTIONS";
        }

        public async Task ExecuteAsync(AnalysisResultItem item)
        {
            _unitService.SetPresentUnitsFromMainWindow();
            EtabsTableResult result = await _tableService.GetTableAsync(item.EtabsTableName);
            _excelExportService.ExportTable(result);
        }
    }
}
