using ExcelCSIToolBox.Application.Interfaces.Etabs.ElementConnectivity;

namespace ExcelCSIToolBox.Infrastructure.Services.Etabs.ElementConnectivity
{
    public class EtabsElementConnectivityServices
    {
        public EtabsElementConnectivityServices(IEtabsElementConnectivityRouter router)
        {
            Router = router;
        }

        public IEtabsElementConnectivityRouter Router { get; private set; }
    }
}
