using System.Collections.Generic;
using ExcelCSIToolBox.Core.Common.Results;
using ExcelCSIToolBox.Data;
using ExcelCSIToolBox.Data.DTOs.CSI;
using ExcelCSIToolBox.Data.Models;

namespace ExcelCSIToolBox.Infrastructure.CSISapModel
{
    internal interface ICSISapModelConnectionAdapter<TSapModel>
    {
        string ProductName { get; }

        OperationResult<IReadOnlyList<CSISapModelRunningInstanceDTO>> GetRunningInstances();

        OperationResult<CSISapModelConnectionInfoDTO> TryAttachToRunningInstance();

        OperationResult<CSISapModelConnectionInfoDTO> AttachToRunningInstance(string instanceId);

        OperationResult<CSISapModelConnectionInfoDTO> GetCurrentConnection();

        OperationResult<TSapModel> EnsureSapModel();

        OperationResult CloseCurrentInstance();
    }
}

