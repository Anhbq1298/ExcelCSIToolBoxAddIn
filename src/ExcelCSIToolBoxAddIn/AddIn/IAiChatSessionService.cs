using System;
using System.Threading;
using System.Threading.Tasks;
using ExcelCSIToolBox.AI.Agent;

namespace ExcelCSIToolBoxAddIn.AddIn
{
    public interface IAiChatSessionService
    {
        string ModelName { get; }

        Task<AiAgentResponse> SendAsync(
            string userMessage,
            Action<string> onAssistantToken,
            CancellationToken cancellationToken);

        bool IsEtabsAttached();
        bool IsSap2000Attached();
    }
}
