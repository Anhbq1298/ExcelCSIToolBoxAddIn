using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using ExcelCSIToolBox.AI.Mcp.Contracts;
using ExcelCSIToolBox.AI.Mcp.Tools;
using ExcelCSIToolBox.Core.Abstractions.CSI;
using ExcelCSIToolBox.Core.Common.Results;
using Newtonsoft.Json;

namespace ExcelCSIToolBox.AI.Mcp.Tools.CSI
{
    /// <summary>
    /// Read-only MCP tool: returns all currently selected objects (points, frames, shells)
    /// from the attached running ETABS or SAP2000 model.
    /// </summary>
    public class CsiGetSelectedObjectsTool : IMcpTool
    {
        private readonly ICsiReadOnlySelectionService _selectionService;

        public CsiGetSelectedObjectsTool(ICsiReadOnlySelectionService selectionService)
        {
            _selectionService = selectionService
                ?? throw new ArgumentNullException(nameof(selectionService));
        }

        public string Name        => "CSI.GetSelectedObjects";
        public string Description => "Returns all currently selected objects (points, frames, shells) from the attached running model.";
        public bool   IsReadOnly  => true;

        public Task<ToolCallResponse> ExecuteAsync(string argumentsJson, CancellationToken cancellationToken)
        {
            OperationResult<List<CsiSelectedObjectDto>> result = _selectionService.GetSelectedObjects();

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

            var payload = new
            {
                Count   = result.Data.Count,
                Objects = result.Data
            };

            return Task.FromResult(new ToolCallResponse
            {
                ToolName   = Name,
                Success    = true,
                Message    = $"Found {result.Data.Count} selected object(s).",
                ResultJson = JsonConvert.SerializeObject(payload)
            });
        }
    }
}
