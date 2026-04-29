using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using ExcelCSIToolBox.AI.Mcp.Contracts;
using ExcelCSIToolBox.AI.Mcp.Tools;
using ExcelCSIToolBox.Core.Abstractions.CSI;
using ExcelCSIToolBox.Core.Common.Results;
using Newtonsoft.Json;

namespace ExcelCSIToolBox.AI.Mcp.Tools.CSI.Frames
{
    /// <summary>
    /// Read-only MCP tool: returns section names assigned to currently selected frame objects.
    /// </summary>
    public class CsiGetSelectedFrameSectionsTool : IMcpTool
    {
        private readonly ICsiReadOnlyFrameService _frameService;

        public CsiGetSelectedFrameSectionsTool(ICsiReadOnlyFrameService frameService)
        {
            _frameService = frameService
                ?? throw new ArgumentNullException(nameof(frameService));
        }

        public string Name        => "CSI.GetSelectedFrameSections";
        public string Description => "Returns section property names assigned to currently selected frame objects.";
        public bool   IsReadOnly  => true;

        public Task<ToolCallResponse> ExecuteAsync(string argumentsJson, CancellationToken cancellationToken)
        {
            OperationResult<List<FrameSectionAssignmentDto>> result = _frameService.GetSelectedFrameSections();

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
                Count       = result.Data.Count,
                Assignments = result.Data
            };

            return Task.FromResult(new ToolCallResponse
            {
                ToolName   = Name,
                Success    = true,
                Message    = $"Retrieved sections for {result.Data.Count} selected frame(s).",
                ResultJson = JsonConvert.SerializeObject(payload)
            });
        }
    }
}
