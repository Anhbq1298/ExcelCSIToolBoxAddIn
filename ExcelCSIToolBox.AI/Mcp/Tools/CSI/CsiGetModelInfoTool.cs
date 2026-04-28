using System;
using System.Threading;
using System.Threading.Tasks;
using ExcelCSIToolBox.AI.Mcp.Contracts;
using ExcelCSIToolBox.AI.Mcp.Tools;
using ExcelCSIToolBox.Core.Abstractions.CSI;
using ExcelCSIToolBox.Core.Common.Results;
using ExcelCSIToolBox.Data;
using ExcelCSIToolBox.Data.DTOs.CSI;
using Newtonsoft.Json;

namespace ExcelCSIToolBox.AI.Mcp.Tools.CSI
{
    /// <summary>
    /// Read-only MCP tool: returns the current model file path and product info
    /// from the attached running ETABS or SAP2000 instance.
    /// </summary>
    public class CsiGetModelInfoTool : IMcpTool
    {
        private readonly ICsiReadOnlyConnectionService _connectionService;

        public CsiGetModelInfoTool(ICsiReadOnlyConnectionService connectionService)
        {
            _connectionService = connectionService
                ?? throw new ArgumentNullException(nameof(connectionService));
        }

        public string Name        => "CSI.GetModelInfo";
        public string Description => "Returns real model file path and product info from the attached running ETABS/SAP2000 model.";
        public bool   IsReadOnly  => true;

        public Task<ToolCallResponse> ExecuteAsync(string argumentsJson, CancellationToken cancellationToken)
        {
            OperationResult<CSISapModelConnectionInfoDTO> result = _connectionService.GetCurrentModelInfo();

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

            // Build a safe payload — never expose raw COM object references.
            var payload = new
            {
                Product     = _connectionService.ProductName,
                IsConnected = result.Data.IsConnected,
                ModelPath   = result.Data.ModelPath,
                ModelFile   = result.Data.ModelFileName,
                CurrentUnit = result.Data.ModelCurrentUnit
            };

            return Task.FromResult(new ToolCallResponse
            {
                ToolName   = Name,
                Success    = true,
                Message    = "Model info retrieved successfully.",
                ResultJson = JsonConvert.SerializeObject(payload)
            });
        }
    }
}
