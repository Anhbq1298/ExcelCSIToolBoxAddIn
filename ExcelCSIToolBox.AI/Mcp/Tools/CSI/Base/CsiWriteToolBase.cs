using System;
using System.Threading;
using System.Threading.Tasks;
using ExcelCSIToolBox.AI.Mcp.Contracts;
using ExcelCSIToolBox.AI.Mcp.Tools;
using ExcelCSIToolBox.Core.Abstractions.CSI;
using ExcelCSIToolBox.Core.Common.Results;
using ExcelCSIToolBox.Core.Models.CSI;
using Newtonsoft.Json;

namespace ExcelCSIToolBox.AI.Mcp.Tools.CSI.Base
{
    public abstract class CsiWriteToolBase<TArgs> : IMcpTool, IMcpToolMetadata
        where TArgs : class, new()
    {
        protected readonly ICsiModelCommandService CommandService;

        protected CsiWriteToolBase(ICsiModelCommandService commandService)
        {
            CommandService = commandService ?? throw new ArgumentNullException(nameof(commandService));
        }

        public abstract string Name { get; }
        public abstract string Description { get; }
        public abstract string Title { get; }
        public abstract string Category { get; }
        public abstract string SubCategory { get; }
        public abstract CsiMethodRiskLevel RiskLevel { get; }
        public abstract bool RequiresConfirmation { get; }
        public abstract bool SupportsDryRun { get; }
        public bool IsReadOnly => false;

        public Task<ToolCallResponse> ExecuteAsync(string argumentsJson, CancellationToken cancellationToken)
        {
            TArgs args;
            try
            {
                args = JsonConvert.DeserializeObject<TArgs>(argumentsJson ?? "{}") ?? new TArgs();
            }
            catch (Exception ex)
            {
                return Task.FromResult(Fail("Invalid arguments: " + ex.Message));
            }

            try
            {
                return Task.FromResult(Execute(args));
            }
            catch (Exception ex)
            {
                return Task.FromResult(Fail("Tool execution failed: " + ex.Message));
            }
        }

        protected abstract ToolCallResponse Execute(TArgs args);

        protected ToolCallResponse Preview(CsiWritePreview preview)
        {
            return new ToolCallResponse
            {
                ToolName = Name,
                Success = true,
                Message = preview.Summary,
                ResultJson = JsonConvert.SerializeObject(preview)
            };
        }

        protected ToolCallResponse Result(OperationResult result)
        {
            return new ToolCallResponse
            {
                ToolName = Name,
                Success = result.IsSuccess,
                Message = result.Message,
                ResultJson = JsonConvert.SerializeObject(new
                {
                    result.IsSuccess,
                    result.Message
                })
            };
        }

        protected ToolCallResponse Fail(string message)
        {
            return new ToolCallResponse
            {
                ToolName = Name,
                Success = false,
                Message = message,
                ResultJson = null
            };
        }
    }

    public class DryRunConfirmedArgs
    {
        public bool DryRun { get; set; } = false;
        public bool Confirmed { get; set; }
    }

    public class LowRiskWriteArgs : DryRunConfirmedArgs
    {
        public LowRiskWriteArgs()
        {
            DryRun = false;
            Confirmed = true;
        }
    }
}
