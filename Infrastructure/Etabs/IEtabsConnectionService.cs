using System.Collections.Generic;
using ExcelCSIToolBoxAddIn.Common.Results;

namespace ExcelCSIToolBoxAddIn.Infrastructure.Etabs
{
    public interface IEtabsConnectionService
    {
        OperationResult<EtabsConnectionInfo> TryAttachToRunningInstance();

        OperationResult<EtabsConnectionInfo> GetCurrentConnection();

        OperationResult<IReadOnlyList<EtabsPointData>> GetSelectedPointsFromActiveModel();
    }
}
