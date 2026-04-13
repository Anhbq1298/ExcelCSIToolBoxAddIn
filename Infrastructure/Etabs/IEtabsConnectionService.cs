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

        OperationResult<EtabsAddPointsResult> AddPointsCartesian(IReadOnlyList<EtabsPointCartesianInput> points);

        OperationResult<IReadOnlyList<EtabsPointData>> GetSelectedPointsFromActiveModel();
    }
}
