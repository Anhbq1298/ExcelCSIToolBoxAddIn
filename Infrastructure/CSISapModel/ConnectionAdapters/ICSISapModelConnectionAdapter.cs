using ExcelCSIToolBoxAddIn.Common.Results;
using ExcelCSIToolBoxAddIn.Data;
using ExcelCSIToolBoxAddIn.Data.DTOs;
using ExcelCSIToolBoxAddIn.Data.Models;

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
