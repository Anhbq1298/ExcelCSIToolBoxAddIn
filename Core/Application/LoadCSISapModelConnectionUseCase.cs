using ExcelCSIToolBoxAddIn.Common.Results;
using ExcelCSIToolBoxAddIn.Infrastructure.CSISapModel;
using ExcelCSIToolBoxAddIn.Data;
using ExcelCSIToolBoxAddIn.Data.DTOs;
using ExcelCSIToolBoxAddIn.Data.Models;


namespace ExcelCSIToolBoxAddIn.Core.Application
{
    /// <summary>
    /// Application use case for loading CSI connection state for the toolbox shell.
    /// </summary>
    public class LoadCSISapModelConnectionUseCase
    {
        private readonly ICSISapModelConnectionService _connectionService;

        public LoadCSISapModelConnectionUseCase(ICSISapModelConnectionService connectionService)
        {
            _connectionService = connectionService;
        }

        public OperationResult<CSISapModelConnectionInfoDTO> Execute()
        {
            return _connectionService.TryAttachToRunningInstance();
        }
    }
}
