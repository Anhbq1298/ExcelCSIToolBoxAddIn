using ExcelCSIToolBox.Application.Interfaces.Etabs.ElementConnectivity;
using ExcelCSIToolBox.Application.Features.Connectivity;

namespace ExcelCSIToolBox.Infrastructure.Services.Etabs.ElementConnectivity
{
    public class EtabsElementConnectivityServices
    {
        public EtabsElementConnectivityServices(
            IEtabsElementConnectivityRouter router,
            ExportSelectedObjectConnectivityUseCase exportSelectedObjectConnectivity)
        {
            Router = router;
            ExportSelectedObjectConnectivity = exportSelectedObjectConnectivity;
        }

        public IEtabsElementConnectivityRouter Router { get; private set; }

        public ExportSelectedObjectConnectivityUseCase ExportSelectedObjectConnectivity { get; private set; }
    }
}
