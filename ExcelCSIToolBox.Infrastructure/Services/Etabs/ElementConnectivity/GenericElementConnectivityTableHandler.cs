using System.Collections.Generic;
using System.Threading.Tasks;
using ExcelCSIToolBox.Application.Interfaces.Etabs;
using ExcelCSIToolBox.Application.Interfaces.Etabs.ElementConnectivity;
using ExcelCSIToolBox.Application.Interfaces.Excel;
using ExcelCSIToolBox.Core.Models.ElementConnectivity;
using ExcelCSIToolBox.Core.Models.EtabsTables;

namespace ExcelCSIToolBox.Infrastructure.Services.Etabs.ElementConnectivity
{
    public class GenericElementConnectivityTableHandler : IEtabsElementConnectivityHandler
    {
        private readonly HashSet<string> _supportedKeys;
        private readonly IEtabsDatabaseTableService _tableService;
        private readonly IExcelExportService _excelExportService;
        private readonly IEtabsUnitService _unitService;

        public GenericElementConnectivityTableHandler(
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

        public async Task ExecuteAsync(ElementConnectivityItem item)
        {
            _unitService.SetPresentUnitsFromMainWindow();
            EtabsTableResult result = await _tableService.GetTableAsync(item.EtabsTableName);
            _excelExportService.ExportTable(result);
        }
    }
}
