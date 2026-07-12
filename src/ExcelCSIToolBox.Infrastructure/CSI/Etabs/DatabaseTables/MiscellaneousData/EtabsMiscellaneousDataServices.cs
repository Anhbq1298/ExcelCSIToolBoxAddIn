using ExcelCSIToolBox.Application.Interfaces.Etabs.MiscellaneousData;

namespace ExcelCSIToolBox.Infrastructure.CSI.Etabs.DatabaseTables.MiscellaneousData
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
