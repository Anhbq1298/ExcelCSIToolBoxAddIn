using ExcelCSIToolBox.Core.Common.Results;
using ExcelCSIToolBox.Core.Abstractions.CSI;
using ExcelCSIToolBox.Data;
using ExcelCSIToolBox.Data.DTOs.CSI;
using ExcelCSIToolBox.Data.Models;


namespace ExcelCSIToolBox.Application.UseCases
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

