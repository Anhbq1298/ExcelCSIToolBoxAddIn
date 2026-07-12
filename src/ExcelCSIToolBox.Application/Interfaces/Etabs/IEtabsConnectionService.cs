using ExcelCSIToolBox.Core.Common.Results;
using ExcelCSIToolBox.Core.Contracts.CSI;

namespace ExcelCSIToolBox.Application.Interfaces.Etabs
{
    public interface IEtabsConnectionService
    {
        object SapModel { get; }

        OperationResult<CSISapModelPresentUnitSystemDTO> GetPresentUnitSystem();

        OperationResult SetPresentUnitSystem(CSISapModelPresentUnitSystemDTO unitSystem);
    }
}
