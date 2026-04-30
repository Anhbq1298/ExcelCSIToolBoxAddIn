using System;
using System.Threading;
using System.Threading.Tasks;
using ExcelCSIToolBox.AI.Mcp.Contracts;
using ExcelCSIToolBox.AI.Mcp.Tools;
using ExcelCSIToolBox.Core.Abstractions.CSI;
using ExcelCSIToolBox.Core.Common.Results;
using ExcelCSIToolBox.Core.Models.CSI;
using ExcelCSIToolBox.Data.CSISapModel.Random;
using ExcelCSIToolBox.Data.DTOs.CSI;
using Newtonsoft.Json;

namespace ExcelCSIToolBox.AI.Mcp.Tools.CSI.Random
{
    public sealed class RandomGenerateObjectsTool : IMcpTool, IMcpToolMetadata
    {
        private readonly ICSISapModelConnectionService _etabsService;
        private readonly ICSISapModelConnectionService _sap2000Service;
        private readonly ICsiRandomObjectGenerationService _randomService;

        public RandomGenerateObjectsTool(
            ICSISapModelConnectionService etabsService,
            ICSISapModelConnectionService sap2000Service,
            ICsiRandomObjectGenerationService randomService)
        {
            _etabsService = etabsService ?? throw new ArgumentNullException(nameof(etabsService));
            _sap2000Service = sap2000Service ?? throw new ArgumentNullException(nameof(sap2000Service));
            _randomService = randomService ?? throw new ArgumentNullException(nameof(randomService));
        }

        public string Name => "random.generate_objects";
        public string Title => "Generate Random CSI Objects";
        public string Category => "Random";
        public string SubCategory => "Creation";
        public string Description => "Generates bounded random points, frames, and shell/area objects using safe defaults when count or bounds are omitted.";
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

                RandomCsiObjectRequestDto request = JsonConvert.DeserializeObject<RandomCsiObjectRequestDto>(argumentsJson ?? "{}")
                    ?? new RandomCsiObjectRequestDto();
                OperationResult<RandomCsiObjectResultDto> result = _randomService.Generate(serviceResult.Data, request);
                if (!result.IsSuccess)
                {
                    return Task.FromResult(Fail(result.Message));
                }

                return Task.FromResult(new ToolCallResponse
                {
                    ToolName = Name,
                    Success = result.Data.FailedItems == 0,
                    Message = result.Data.FailedItems == 0 ? "Success" : "Failure",
                    ResultJson = JsonConvert.SerializeObject(result.Data)
                });
            }
            catch (Exception ex)
            {
                return Task.FromResult(Fail("Random object generation failed: " + ex.Message));
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
