using System;
using System.Threading;
using System.Threading.Tasks;
using ExcelCSIToolBox.AI.Mcp.Contracts;
using ExcelCSIToolBox.AI.Mcp.Tools;
using ExcelCSIToolBox.Core.Abstractions.CSI;
using ExcelCSIToolBox.Core.Common.Results;
using ExcelCSIToolBox.Core.Models.CSI;
using ExcelCSIToolBox.Data.CSISapModel.Truss;
using ExcelCSIToolBox.Data.DTOs.CSI;
using ExcelCSIToolBox.Infrastructure.CSISapModel.Truss;
using Newtonsoft.Json;

namespace ExcelCSIToolBox.AI.Mcp.Tools.CSI.Truss
{
    public sealed class TrussGenerateHoweTool : IMcpTool, IMcpToolMetadata
    {
        private readonly ICSISapModelConnectionService _etabsService;
        private readonly ICSISapModelConnectionService _sap2000Service;
        private readonly CsiHoweTrussGenerationService _trussService;

        public TrussGenerateHoweTool(
            ICSISapModelConnectionService etabsService,
            ICSISapModelConnectionService sap2000Service)
        {
            _etabsService = etabsService ?? throw new ArgumentNullException(nameof(etabsService));
            _sap2000Service = sap2000Service ?? throw new ArgumentNullException(nameof(sap2000Service));
            _trussService = new CsiHoweTrussGenerationService();
        }

        public string Name => "truss.generate_howe";
        public string Title => "Generate Howe Truss";
        public string Category => "Truss";
        public string SubCategory => "Generation";
        public string Description => "Generates a symmetric Howe truss. Chords are continuous frame segments; vertical and brace members are released at both ends.";
        public bool IsReadOnly => false;
        public CsiMethodRiskLevel RiskLevel => CsiMethodRiskLevel.Low;
        public bool RequiresConfirmation => false;
        public bool SupportsDryRun => false;

        public Task<ToolCallResponse> ExecuteAsync(string argumentsJson, CancellationToken cancellationToken)
        {
            try
            {
                OperationResult<ICSISapModelConnectionService> serviceResult = GetActiveService();
                if (!serviceResult.IsSuccess)
                {
                    return Task.FromResult(Fail(serviceResult.Message));
                }

                HoweTrussRequestDto request = JsonConvert.DeserializeObject<HoweTrussRequestDto>(argumentsJson ?? "{}")
                    ?? new HoweTrussRequestDto();
                OperationResult<HoweTrussResultDto> result = _trussService.Generate(serviceResult.Data, request);
                if (!result.IsSuccess)
                {
                    return Task.FromResult(Fail(result.Message));
                }

                return Task.FromResult(new ToolCallResponse
                {
                    ToolName = Name,
                    Success = result.Data.Success,
                    Message = result.Data.Success ? "Success" : "Failure",
                    ResultJson = JsonConvert.SerializeObject(result.Data)
                });
            }
            catch (Exception ex)
            {
                return Task.FromResult(Fail("Howe truss generation failed: " + ex.Message));
            }
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

            return OperationResult<ICSISapModelConnectionService>.Failure("active CSI model is not available.");
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
