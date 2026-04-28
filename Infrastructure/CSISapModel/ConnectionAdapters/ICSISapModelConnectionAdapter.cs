using ExcelCSIToolBoxAddIn.Common.Results;

namespace ExcelCSIToolBoxAddIn.Infrastructure.CSISapModel
{
    internal interface ICSISapModelConnectionAdapter<TSapModel>
    {
        string ProductName { get; }

        OperationResult<CSISapModelConnectionInfoDTO> TryAttachToRunningInstance();

        OperationResult<CSISapModelConnectionInfoDTO> GetCurrentConnection();

        OperationResult<TSapModel> EnsureSapModel();

        OperationResult CloseCurrentInstance();
    }
}
