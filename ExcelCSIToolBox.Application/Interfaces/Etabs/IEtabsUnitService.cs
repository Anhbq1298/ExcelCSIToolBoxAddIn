using ExcelCSIToolBox.Data.DTOs.CSI;

namespace ExcelCSIToolBox.Application.Interfaces.Etabs
{
    public interface IEtabsUnitService
    {
        void SetSelectedUnitSystem(
            CSISapModelPresentUnitSystemDTO unitSystem,
            string displayName,
            string presentUnitsText);

        void SetPresentUnitsFromMainWindow();
    }
}
