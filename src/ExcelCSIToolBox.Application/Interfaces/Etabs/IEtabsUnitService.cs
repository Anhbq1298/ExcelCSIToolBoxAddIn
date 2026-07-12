using ExcelCSIToolBox.Application.Services;
using ExcelCSIToolBox.Core.Common.Results;
using ExcelCSIToolBox.Core.Contracts.CSI;

namespace ExcelCSIToolBox.Application.Interfaces.Etabs
{
    public interface IEtabsUnitService
    {
        void SetSelectedUnitSystem(
            CSISapModelPresentUnitSystemDTO unitSystem,
            string displayName,
            string presentUnitsText);

        void SetPresentUnitsFromMainWindow();

        OperationResult<CsiPresentUnitScope> CreatePresentUnitScopeFromMainWindow();
    }
}
