using ExcelCSIToolBox.Application.Interfaces.Etabs.MiscellaneousData;

namespace ExcelCSIToolBox.Infrastructure.Services.Etabs.MiscellaneousData
{
    public class EtabsMiscellaneousDataServices
    {
        public EtabsMiscellaneousDataServices(IEtabsMiscellaneousDataRouter router)
        {
            Router = router;
        }

        public IEtabsMiscellaneousDataRouter Router { get; private set; }
    }
}
