using ExcelCSIToolBox.Core.Common.Results;
using ExcelCSIToolBox.Data.DTOs.CSI;

namespace ExcelCSIToolBox.Application.Interfaces.Etabs
{
    public interface IEtabsConnectionService
    {
        object SapModel { get; }

        OperationResult SetPresentUnitSystem(CSISapModelPresentUnitSystemDTO unitSystem);
    }
}
