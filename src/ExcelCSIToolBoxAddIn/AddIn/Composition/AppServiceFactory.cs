using System.Collections.Generic;
using ExcelCSIToolBox.Application.Interfaces.Etabs;
using ExcelCSIToolBox.Application.Interfaces.Etabs.AnalysisResults;
using ExcelCSIToolBox.Application.Interfaces.Etabs.ElementConnectivity;
using ExcelCSIToolBox.Application.Interfaces.Etabs.MiscellaneousData;
using ExcelCSIToolBox.Application.Interfaces.Excel;
using ExcelCSIToolBox.Application.Features.Connectivity;
using ExcelCSIToolBox.Core.Abstractions.CSI;
using ExcelCSIToolBox.Infrastructure.CSI.Common;
using ExcelCSIToolBox.Infrastructure.CSI.Etabs.DatabaseTables;
using ExcelCSIToolBox.Infrastructure.CSI.Etabs.AnalysisResults;
using ExcelCSIToolBox.Infrastructure.CSI.Etabs.AnalysisResults.JointOutput;
using ExcelCSIToolBox.Infrastructure.CSI.Etabs.AnalysisResults.StructureOutput;
using ExcelCSIToolBox.Infrastructure.CSI.Etabs.Connectivity;
using ExcelCSIToolBox.Infrastructure.CSI.Etabs.DatabaseTables.MiscellaneousData;
using ExcelCSIToolBox.Infrastructure.CSI.Etabs.Loadings.ShellUniformLoadSets;
using ExcelCSIToolBox.Infrastructure.CSI.Etabs.Units;
using ExcelCSIToolBox.Infrastructure.CSI.Etabs.Selection;
using ExcelCSIToolBox.Infrastructure.Excel.Interop.Writing;

namespace ExcelCSIToolBoxAddIn.AddIn.Composition
{
    public static class AppServiceFactory
    {
        public static readonly ICsiApiDispatcher CsiApiDispatcher = new CurrentThreadCsiApiDispatcher();

        public static EtabsAnalysisResultServices CreateAnalysisResultServices(
            ICSISapModelConnectionService csiConnectionService,
            IEtabsUnitService unitService = null)
        {
            IEtabsConnectionService connectionService = new EtabsConnectionService(csiConnectionService);
            unitService = unitService ?? new EtabsUnitService(connectionService);
            IEtabsDatabaseTableService tableService = new EtabsDatabaseTableService(connectionService, CsiApiDispatcher);
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

        public static EtabsMiscellaneousDataServices CreateMiscellaneousDataServices(
            ICSISapModelConnectionService csiConnectionService,
            IEtabsUnitService unitService = null)
        {
            IEtabsConnectionService connectionService = new EtabsConnectionService(csiConnectionService);
            unitService = unitService ?? new EtabsUnitService(connectionService);
            IEtabsDatabaseTableService tableService = new EtabsDatabaseTableService(connectionService, CsiApiDispatcher);
            IExcelExportService excelService = new ExcelExportService();

            List<IEtabsMiscellaneousDataHandler> handlers = new List<IEtabsMiscellaneousDataHandler>
            {
                new GenericMiscellaneousDataTableHandler(
                    EtabsMiscellaneousDataRegistry.GetSupportedKeysForGenericTableExport(),
                    tableService,
                    excelService,
                    unitService)
            };

            IEtabsMiscellaneousDataRouter router = new EtabsMiscellaneousDataRouter(handlers);
            return new EtabsMiscellaneousDataServices(router);
        }

        public static EtabsElementConnectivityServices CreateElementConnectivityServices(
            ICSISapModelConnectionService csiConnectionService,
            IEtabsUnitService unitService = null)
        {
            IEtabsConnectionService connectionService = new EtabsConnectionService(csiConnectionService);
            unitService = unitService ?? new EtabsUnitService(connectionService);
            IEtabsDatabaseTableService tableService = new EtabsDatabaseTableService(connectionService, CsiApiDispatcher);
            IExcelExportService excelService = new ExcelExportService();

            List<IEtabsElementConnectivityHandler> handlers = new List<IEtabsElementConnectivityHandler>
            {
                new GenericElementConnectivityTableHandler(
                    EtabsElementConnectivityRegistry.GetSupportedKeysForGenericTableExport(),
                    tableService,
                    excelService,
                    unitService)
            };

            IEtabsElementConnectivityRouter router = new EtabsElementConnectivityRouter(handlers);
            var identityResolver = new EtabsSelectedObjectIdentityResolver(connectionService);
            var exportSelectedObjectConnectivity = new ExportSelectedObjectConnectivityUseCase(tableService, identityResolver);
            return new EtabsElementConnectivityServices(router, exportSelectedObjectConnectivity);
        }

        public static IEtabsShellUniformLoadSetSelectionService CreateShellUniformLoadSetSelectionService(
            ICSISapModelConnectionService csiConnectionService)
        {
            IEtabsConnectionService connectionService = new EtabsConnectionService(csiConnectionService);
            return new EtabsShellUniformLoadSetSelectionService(connectionService, CsiApiDispatcher);
        }
    }
}
