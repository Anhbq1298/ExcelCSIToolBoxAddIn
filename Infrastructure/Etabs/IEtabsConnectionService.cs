using System.Collections.Generic;
using ExcelCSIToolBoxAddIn.Common.Results;

namespace ExcelCSIToolBoxAddIn.Infrastructure.Etabs
{
    public interface IEtabsConnectionService
    {
        OperationResult<EtabsConnectionInfo> TryAttachToRunningInstance();

        OperationResult<EtabsConnectionInfo> GetCurrentConnection();

        OperationResult CloseCurrentEtabsInstance();

        OperationResult SelectPointsByUniqueNames(IReadOnlyList<string> uniqueNames);

        // pointInputs must be executed exactly in the given order.
        // Duplicate rows are valid and must not be merged or de-duplicated.
        OperationResult<EtabsAddPointsResult> AddPointsByCartesian(IReadOnlyList<EtabsPointCartesianInput> pointInputs);

        OperationResult<IReadOnlyList<EtabsPointData>> GetSelectedPointsFromActiveModel();

        OperationResult<string> GetCurrentModelUnitsDisplayText();
    }
}
