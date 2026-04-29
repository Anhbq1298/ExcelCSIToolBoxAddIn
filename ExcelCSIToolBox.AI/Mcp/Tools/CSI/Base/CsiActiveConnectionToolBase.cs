using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using ExcelCSIToolBox.AI.Mcp.Contracts;
using ExcelCSIToolBox.Core.Abstractions.CSI;
using ExcelCSIToolBox.Core.Common.Results;
using ExcelCSIToolBox.Core.Models.CSI;
using ExcelCSIToolBox.Data.DTOs.CSI;
using Newtonsoft.Json;

namespace ExcelCSIToolBox.AI.Mcp.Tools.CSI.Base
{
    public abstract class CsiActiveConnectionToolBase : IMcpTool, IMcpToolMetadata
    {
        private readonly ICSISapModelConnectionService _etabsService;
        private readonly ICSISapModelConnectionService _sap2000Service;

        protected CsiActiveConnectionToolBase(
            ICSISapModelConnectionService etabsService,
            ICSISapModelConnectionService sap2000Service)
        {
            _etabsService = etabsService ?? throw new ArgumentNullException(nameof(etabsService));
            _sap2000Service = sap2000Service ?? throw new ArgumentNullException(nameof(sap2000Service));
        }

        public abstract string Name { get; }
        public abstract string Title { get; }
        public abstract string Category { get; }
        public abstract string SubCategory { get; }
        public abstract string Description { get; }
        public abstract bool IsReadOnly { get; }
        public abstract CsiMethodRiskLevel RiskLevel { get; }
        public abstract bool RequiresConfirmation { get; }
        public abstract bool SupportsDryRun { get; }

        public Task<ToolCallResponse> ExecuteAsync(string argumentsJson, CancellationToken cancellationToken)
        {
            try
            {
                OperationResult<ICSISapModelConnectionService> serviceResult = GetActiveService();
                if (!serviceResult.IsSuccess)
                {
                    return Task.FromResult(Fail(serviceResult.Message));
                }

                return Task.FromResult(Execute(serviceResult.Data, argumentsJson ?? "{}"));
            }
            catch (Exception ex)
            {
                return Task.FromResult(Fail("CSI tool execution failed: " + ex.Message));
            }
        }

        protected abstract ToolCallResponse Execute(ICSISapModelConnectionService service, string argumentsJson);

        protected TArgs ReadArgs<TArgs>(string argumentsJson) where TArgs : class, new()
        {
            return JsonConvert.DeserializeObject<TArgs>(argumentsJson ?? "{}") ?? new TArgs();
        }

        protected ToolCallResponse Result<T>(OperationResult<T> result)
        {
            return result.IsSuccess
                ? Ok(result.Message ?? "CSI query completed.", result.Data)
                : Fail(result.Message);
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

        protected ToolCallResponse Preview(CsiWritePreview preview)
        {
            return Ok(preview.Summary, preview);
        }

        protected ToolCallResponse Ok(string message)
        {
            return Ok(message, null);
        }

        protected ToolCallResponse Ok(string message, object payload)
        {
            return new ToolCallResponse
            {
                ToolName = Name,
                Success = true,
                Message = message,
                ResultJson = JsonConvert.SerializeObject(payload)
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

        protected static int Count<T>(IReadOnlyList<T> items)
        {
            return items == null ? 0 : items.Count;
        }

        private OperationResult<ICSISapModelConnectionService> GetActiveService()
        {
            OperationResult<CSISapModelConnectionInfoDTO> etabs = _etabsService.GetCurrentConnection();
            if (IsConnected(etabs))
            {
                return OperationResult<ICSISapModelConnectionService>.Success(_etabsService);
            }

            OperationResult<CSISapModelConnectionInfoDTO> sap2000 = _sap2000Service.GetCurrentConnection();
            if (IsConnected(sap2000))
            {
                return OperationResult<ICSISapModelConnectionService>.Success(_sap2000Service);
            }

            etabs = _etabsService.TryAttachToRunningInstance();
            if (IsConnected(etabs))
            {
                return OperationResult<ICSISapModelConnectionService>.Success(_etabsService);
            }

            sap2000 = _sap2000Service.TryAttachToRunningInstance();
            if (IsConnected(sap2000))
            {
                return OperationResult<ICSISapModelConnectionService>.Success(_sap2000Service);
            }

            return OperationResult<ICSISapModelConnectionService>.Failure("No ETABS or SAP2000 model is attached.");
        }

        private static bool IsConnected(OperationResult<CSISapModelConnectionInfoDTO> result)
        {
            return result != null &&
                   result.IsSuccess &&
                   result.Data != null &&
                   result.Data.IsConnected &&
                   result.Data.SapModel != null;
        }
    }

    public sealed class NamesDryRunArgs : DryRunConfirmedArgs
    {
        public List<string> Names { get; set; }
    }
}
