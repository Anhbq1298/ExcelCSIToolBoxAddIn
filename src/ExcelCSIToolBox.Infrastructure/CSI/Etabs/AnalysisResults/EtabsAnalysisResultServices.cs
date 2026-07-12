using ExcelCSIToolBox.Application.Interfaces.Etabs;
using ExcelCSIToolBox.Application.Interfaces.Etabs.AnalysisResults;

namespace ExcelCSIToolBox.Infrastructure.CSI.Etabs.AnalysisResults
{
    public class EtabsAnalysisResultServices
    {
        public EtabsAnalysisResultServices(
            IEtabsAnalysisResultRouter router,
            IEtabsUnitService unitService)
        {
            Router = router;
            UnitService = unitService;
        }

        public IEtabsAnalysisResultRouter Router { get; private set; }

        public IEtabsUnitService UnitService { get; private set; }
    }
}
