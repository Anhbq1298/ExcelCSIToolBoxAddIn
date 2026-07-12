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

namespace ExcelCSIToolBox.AI.Mcp.Tools.CSI.Shells
{
    /// <summary>
    /// Read-only MCP tool: returns shell/area object names from the active CSI model.
    /// </summary>
    public sealed class CsiGetShellNamesTool : IMcpTool, IMcpToolMetadata
    {
        private readonly ICSISapModelConnectionService _etabsService;
        private readonly ICSISapModelConnectionService _sap2000Service;

        public CsiGetShellNamesTool(
            ICSISapModelConnectionService etabsService,
            ICSISapModelConnectionService sap2000Service)
        {
            _etabsService = etabsService ?? throw new ArgumentNullException(nameof(etabsService));
            _sap2000Service = sap2000Service ?? throw new ArgumentNullException(nameof(sap2000Service));
        }

        public string Name => "shells.get_all_names";
        public string Title => "Get Shell / Area Names";
        public string Category => "Shells / Areas";
        public string SubCategory => "Read";
        public string Description => "Returns all shell/area object names from the attached ETABS or SAP2000 model.";
        public bool IsReadOnly => true;
        public CsiMethodRiskLevel RiskLevel => CsiMethodRiskLevel.None;
        public bool RequiresConfirmation => false;
        public bool SupportsDryRun => false;

        public Task<ToolCallResponse> ExecuteAsync(string argumentsJson, CancellationToken cancellationToken)
        {
            OperationResult<ICSISapModelConnectionService> serviceResult = GetActiveService();
            if (!serviceResult.IsSuccess)
            {
                return Task.FromResult(Fail(serviceResult.Message));
            }

            OperationResult<IReadOnlyList<string>> namesResult = serviceResult.Data.GetShellNames();
            if (!namesResult.IsSuccess)
            {
                return Task.FromResult(Fail(namesResult.Message));
            }

            IReadOnlyList<string> names = namesResult.Data ?? new string[0];
            var payload = new
            {
                Product = serviceResult.Data.ProductName,
                Count = names.Count,
                ShellNames = names
            };

            return Task.FromResult(new ToolCallResponse
            {
                ToolName = Name,
                Success = true,
                Message = $"Found {names.Count} shell/area object(s) in {serviceResult.Data.ProductName}.",
                ResultJson = JsonConvert.SerializeObject(payload)
            });
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

        private ToolCallResponse Fail(string message)
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
}
