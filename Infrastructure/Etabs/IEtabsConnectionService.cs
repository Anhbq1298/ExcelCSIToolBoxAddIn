using ExcelCSIToolBoxAddIn.Common.Results;

namespace ExcelCSIToolBoxAddIn.Infrastructure.Etabs
{
    public interface IEtabsConnectionService
    {
        OperationResult<EtabsConnectionInfo> TryAttachToRunningInstance();
    }
}
