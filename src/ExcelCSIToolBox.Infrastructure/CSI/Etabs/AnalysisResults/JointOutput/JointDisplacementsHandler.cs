using System.Threading.Tasks;
using ExcelCSIToolBox.Application.Interfaces.Etabs;
using ExcelCSIToolBox.Application.Interfaces.Etabs.AnalysisResults;
using ExcelCSIToolBox.Application.Interfaces.Excel;
using ExcelCSIToolBox.Core.Models.AnalysisResults;
using ExcelCSIToolBox.Core.Models.EtabsTables;
using ExcelCSIToolBox.Infrastructure.CSI.Etabs.DatabaseTables;
using ExcelCSIToolBox.Infrastructure.CSI.Etabs.Units;

namespace ExcelCSIToolBox.Infrastructure.CSI.Etabs.AnalysisResults.JointOutput
{
    public class JointDisplacementsHandler : IEtabsAnalysisResultHandler
    {
        private readonly IEtabsDatabaseTableService _tableService;
        private readonly IExcelExportService _excelExportService;
        private readonly IEtabsUnitService _unitService;

        public JointDisplacementsHandler(
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
            return key == "JOINT_DISPLACEMENTS";
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