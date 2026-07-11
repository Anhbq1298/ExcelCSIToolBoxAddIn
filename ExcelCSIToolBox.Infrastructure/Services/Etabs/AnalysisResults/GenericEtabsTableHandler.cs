using System.Collections.Generic;
using System.Threading.Tasks;
using ExcelCSIToolBox.Application.Interfaces.Etabs;
using ExcelCSIToolBox.Application.Interfaces.Etabs.AnalysisResults;
using ExcelCSIToolBox.Application.Interfaces.Excel;
using ExcelCSIToolBox.Core.Models.AnalysisResults;
using ExcelCSIToolBox.Core.Models.EtabsTables;
using ExcelCSIToolBox.Infrastructure.Services.Etabs;

namespace ExcelCSIToolBox.Infrastructure.Services.Etabs.AnalysisResults
{
    public class GenericEtabsTableHandler : IEtabsAnalysisResultHandler
    {
        private readonly HashSet<string> _supportedKeys;
        private readonly IEtabsDatabaseTableService _tableService;
        private readonly IExcelExportService _excelExportService;
        private readonly IEtabsUnitService _unitService;

        public GenericEtabsTableHandler(
            IEnumerable<string> supportedKeys,
            IEtabsDatabaseTableService tableService,
            IExcelExportService excelExportService,
            IEtabsUnitService unitService)
        {
            _supportedKeys = new HashSet<string>(supportedKeys ?? new string[0]);
            _tableService = tableService;
            _excelExportService = excelExportService;
            _unitService = unitService;
        }

        public bool CanHandle(string key)
        {
            return key != null && _supportedKeys.Contains(key);
        }

        public async Task ExecuteAsync(AnalysisResultItem item)
        {
            var unitScope = EtabsPresentUnitScopeRunner.Begin(_unitService);
            try
            {
                EtabsTableResult result = await _tableService.GetTableAsync(item.EtabsTableName);
                _excelExportService.ExportTable(result);
            }
            finally
            {
                EtabsPresentUnitScopeRunner.Restore(unitScope, item == null ? null : item.Title);
            }
        }
    }
}
