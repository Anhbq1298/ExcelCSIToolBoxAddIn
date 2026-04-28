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
    /// Read-only MCP tool: returns the unique names of currently selected frame objects.
    /// </summary>
    public class CsiGetSelectedFramesTool : IMcpTool
    {
        private readonly ICsiReadOnlySelectionService _selectionService;

        public CsiGetSelectedFramesTool(ICsiReadOnlySelectionService selectionService)
        {
            _selectionService = selectionService
                ?? throw new ArgumentNullException(nameof(selectionService));
        }

        public string Name        => "CSI.GetSelectedFrames";
        public string Description => "Returns unique names of currently selected frame objects from the attached running model.";
        public bool   IsReadOnly  => true;

        public Task<ToolCallResponse> ExecuteAsync(string argumentsJson, CancellationToken cancellationToken)
        {
            OperationResult<List<string>> result = _selectionService.GetSelectedFrameNames();

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
                Count      = result.Data.Count,
                FrameNames = result.Data
            };

            return Task.FromResult(new ToolCallResponse
            {
                ToolName   = Name,
                Success    = true,
                Message    = $"Found {result.Data.Count} selected frame(s).",
                ResultJson = JsonConvert.SerializeObject(payload)
            });
        }
    }
}
