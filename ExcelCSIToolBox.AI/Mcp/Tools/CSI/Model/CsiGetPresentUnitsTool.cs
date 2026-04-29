using System;
using System.Threading;
using System.Threading.Tasks;
using ExcelCSIToolBox.AI.Mcp.Contracts;
using ExcelCSIToolBox.AI.Mcp.Tools;
using ExcelCSIToolBox.Core.Abstractions.CSI;
using ExcelCSIToolBox.Core.Common.Results;
using Newtonsoft.Json;

namespace ExcelCSIToolBox.AI.Mcp.Tools.CSI.Model
{
    /// <summary>
    /// Read-only MCP tool: returns the current units from the attached running model.
    /// </summary>
    public class CsiGetPresentUnitsTool : IMcpTool
    {
        private readonly ICsiReadOnlyConnectionService _connectionService;

        public CsiGetPresentUnitsTool(ICsiReadOnlyConnectionService connectionService)
        {
            _connectionService = connectionService
                ?? throw new ArgumentNullException(nameof(connectionService));
        }

        public string Name        => "CSI.GetPresentUnits";
        public string Description => "Returns the current model units from the attached running ETABS/SAP2000 model.";
        public bool   IsReadOnly  => true;

        public Task<ToolCallResponse> ExecuteAsync(string argumentsJson, CancellationToken cancellationToken)
        {
            OperationResult<string> result = _connectionService.GetPresentUnits();

            if (!result.IsSuccess)
            {
                return Task.FromResult(new ToolCallResponse
                {
                    ToolName   = Name,
                    Success    = false,
                    Message    = result.Message,
                    ResultJson = null
                });
            }

            var payload = new { Units = result.Data };

            return Task.FromResult(new ToolCallResponse
            {
                ToolName   = Name,
                Success    = true,
                Message    = "Present units retrieved successfully.",
                ResultJson = JsonConvert.SerializeObject(payload)
            });
        }
    }
}
