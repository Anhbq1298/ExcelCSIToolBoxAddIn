using System.Collections.Generic;
using ExcelCSIToolBox.Application.Interfaces.Etabs;
using ExcelCSIToolBox.Application.Interfaces.Etabs.AnalysisResults;
using ExcelCSIToolBox.Application.Interfaces.Excel;
using ExcelCSIToolBox.Core.Abstractions.CSI;
using ExcelCSIToolBox.Infrastructure.Services.Etabs;
using ExcelCSIToolBox.Infrastructure.Services.Etabs.AnalysisResults;
using ExcelCSIToolBox.Infrastructure.Services.Etabs.AnalysisResults.JointOutput;
using ExcelCSIToolBox.Infrastructure.Services.Etabs.AnalysisResults.StructureOutput;
using ExcelCSIToolBox.Infrastructure.Services.Excel;

namespace ExcelCSIToolBoxAddIn.AddIn.Composition
{
    public static class AppServiceFactory
    {
        public static EtabsAnalysisResultServices CreateAnalysisResultServices(
            ICSISapModelConnectionService csiConnectionService)
        {
            IEtabsConnectionService connectionService = new EtabsConnectionService(csiConnectionService);
            IEtabsUnitService unitService = new EtabsUnitService(connectionService);
            IEtabsDatabaseTableService tableService = new EtabsDatabaseTableService(connectionService);
            IExcelExportService excelService = new ExcelExportService();

            List<IEtabsAnalysisResultHandler> handlers = new List<IEtabsAnalysisResultHandler>
            {
                new JointDisplacementsHandler(tableService, excelService, unitService),
                new BaseReactionsHandler(tableService, excelService, unitService),
                new GenericEtabsTableHandler(
                    EtabsAnalysisResultRegistry.GetSupportedKeysForGenericTableExport(),
                    tableService,
                    excelService,
                    unitService)
            };

            IEtabsAnalysisResultRouter router = new EtabsAnalysisResultRouter(handlers);
            return new EtabsAnalysisResultServices(router, unitService);
        }
    }
}
