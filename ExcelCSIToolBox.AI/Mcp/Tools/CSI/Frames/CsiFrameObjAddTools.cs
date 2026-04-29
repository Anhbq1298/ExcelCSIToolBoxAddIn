using System;
using System.Threading;
using System.Threading.Tasks;
using ExcelCSIToolBox.AI.Mcp.Contracts;
using ExcelCSIToolBox.AI.Mcp.Tools;
using ExcelCSIToolBox.Core.Abstractions.CSI;
using ExcelCSIToolBox.Core.Common.Results;
using ExcelCSIToolBox.Core.Models.CSI;
using ExcelCSIToolBox.Data.CSISapModel.FrameObject;
using ExcelCSIToolBox.Data.DTOs.CSI;
using Newtonsoft.Json;

namespace ExcelCSIToolBox.AI.Mcp.Tools.CSI.Frames
{
    public sealed class FramesAddObjectTool : FrameObjAddToolBase
    {
        public FramesAddObjectTool(ICSISapModelConnectionService etabsService, ICSISapModelConnectionService sap2000Service)
            : base(etabsService, sap2000Service)
        {
        }

        public override string Name => "frames.add_object";
        public override string Title => "Add Frame Object";
        public override string Description => "Adds one CSI frame object. The service chooses AddByPoint or AddByCoord from the supplied fields.";

        protected override ToolCallResponse Execute(ICSISapModelConnectionService service, string argumentsJson)
        {
            FrameAddRequestDto request = ReadArgs<FrameAddRequestDto>(argumentsJson);
            OperationResult<FrameAddBatchResultDto> result = service.AddFrameObjects(new FrameAddBatchRequestDto
            {
                Frames = new System.Collections.Generic.List<FrameAddRequestDto> { request }
            });

            if (!result.IsSuccess)
            {
                return Fail(result.Message);
            }

            FrameAddResultDto item = result.Data.Results.Count == 0
                ? new FrameAddResultDto { Success = false, FailureReason = "Frame definition is required." }
                : result.Data.Results[0];

            return new ToolCallResponse
            {
                ToolName = Name,
                Success = item.Success,
                Message = item.Success ? "Success" : "Failure",
                ResultJson = JsonConvert.SerializeObject(item)
            };
        }
    }

    public sealed class FramesAddObjectsTool : FrameObjAddToolBase
    {
        public FramesAddObjectsTool(ICSISapModelConnectionService etabsService, ICSISapModelConnectionService sap2000Service)
            : base(etabsService, sap2000Service)
        {
        }

        public override string Name => "frames.add_objects";
        public override string Title => "Add Frame Objects";
        public override string Description => "Adds multiple CSI frame objects. Each item is independently executed with AddByPoint or AddByCoord.";

        protected override ToolCallResponse Execute(ICSISapModelConnectionService service, string argumentsJson)
        {
            FrameAddBatchRequestDto request = ReadArgs<FrameAddBatchRequestDto>(argumentsJson);
            OperationResult<FrameAddBatchResultDto> result = service.AddFrameObjects(request);
            if (!result.IsSuccess)
            {
                return Fail(result.Message);
            }

            return new ToolCallResponse
            {
                ToolName = Name,
                Success = result.Data.FailureCount == 0,
                Message = result.Data.FailureCount == 0 ? "Success" : "Failure",
                ResultJson = JsonConvert.SerializeObject(result.Data)
            };
        }
    }

    public abstract class FrameObjAddToolBase : IMcpTool, IMcpToolMetadata
    {
        private readonly ICSISapModelConnectionService _etabsService;
        private readonly ICSISapModelConnectionService _sap2000Service;

        protected FrameObjAddToolBase(
            ICSISapModelConnectionService etabsService,
            ICSISapModelConnectionService sap2000Service)
        {
            _etabsService = etabsService ?? throw new ArgumentNullException(nameof(etabsService));
            _sap2000Service = sap2000Service ?? throw new ArgumentNullException(nameof(sap2000Service));
        }

        public abstract string Name { get; }
        public abstract string Title { get; }
        public string Category => "Frames";
        public string SubCategory => "Creation";
        public abstract string Description { get; }
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

                return Task.FromResult(Execute(serviceResult.Data, argumentsJson ?? "{}"));
            }
            catch (Exception ex)
            {
                return Task.FromResult(Fail(ex.Message));
            }
        }

        protected abstract ToolCallResponse Execute(ICSISapModelConnectionService service, string argumentsJson);

        protected TArgs ReadArgs<TArgs>(string argumentsJson) where TArgs : class, new()
        {
            return JsonConvert.DeserializeObject<TArgs>(argumentsJson ?? "{}") ?? new TArgs();
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
    }
}
