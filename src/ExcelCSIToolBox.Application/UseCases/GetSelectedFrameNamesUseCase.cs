using System.Collections.Generic;
using ExcelCSIToolBox.Core.Abstractions.CSI;
using ExcelCSIToolBox.Core.Common.Results;

namespace ExcelCSIToolBox.Application.UseCases
{
    /// <summary>
    /// Retrieves selected frame object names from a CSI model connection.
    /// </summary>
    public class GetSelectedFrameNamesUseCase
    {
        private readonly ICSISapModelConnectionService _connectionService;

        public GetSelectedFrameNamesUseCase(ICSISapModelConnectionService connectionService)
        {
            _connectionService = connectionService;
        }

        public OperationResult<IReadOnlyList<string>> Execute()
        {
            return _connectionService.GetSelectedFramesFromActiveModel();
        }
    }
}
