using System;
using System.Threading;
using System.Threading.Tasks;
using ExcelCSIToolBox.AI.Agent;
using ExcelCSIToolBox.Core.Abstractions.CSI;
using ExcelCSIToolBox.Core.Common.Results;
using ExcelCSIToolBox.Data.DTOs.CSI;

namespace ExcelCSIToolBoxAddIn.AddIn
{
    internal sealed class AiChatSessionService : IAiChatSessionService
    {
        private readonly AiAgentOrchestrator _orchestrator;
        private readonly ICSISapModelConnectionService _etabsConnectionService;
        private readonly ICSISapModelConnectionService _sap2000ConnectionService;

        public AiChatSessionService(
            AiAgentOrchestrator orchestrator,
            ICSISapModelConnectionService etabsConnectionService,
            ICSISapModelConnectionService sap2000ConnectionService,
            string modelName)
        {
            _orchestrator = orchestrator ?? throw new ArgumentNullException(nameof(orchestrator));
            _etabsConnectionService = etabsConnectionService ?? throw new ArgumentNullException(nameof(etabsConnectionService));
            _sap2000ConnectionService = sap2000ConnectionService ?? throw new ArgumentNullException(nameof(sap2000ConnectionService));
            ModelName = string.IsNullOrWhiteSpace(modelName) ? "Not selected" : modelName;
        }

        public string ModelName { get; }

        public Task<AiAgentResponse> SendAsync(
            string userMessage,
            Action<string> onAssistantToken,
            CancellationToken cancellationToken)
        {
            return _orchestrator.SendAsync(userMessage, onAssistantToken, cancellationToken);
        }

        public bool IsEtabsAttached()
        {
            return IsProductAttached(_etabsConnectionService);
        }

        public bool IsSap2000Attached()
        {
            return IsProductAttached(_sap2000ConnectionService);
        }

        private static bool IsProductAttached(ICSISapModelConnectionService connectionService)
        {
            try
            {
                OperationResult<CSISapModelConnectionInfoDTO> result = connectionService.TryAttachToRunningInstance();
                return result != null &&
                       result.IsSuccess &&
                       result.Data != null &&
                       result.Data.IsConnected &&
                       result.Data.SapModel != null;
            }
            catch
            {
                return false;
            }
        }
    }
}
