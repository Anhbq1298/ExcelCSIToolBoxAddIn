using System.Threading;
using System.Threading.Tasks;
using ExcelCSIToolBox.AI.Mcp.Contracts;

namespace ExcelCSIToolBox.AI.Agent
{
    public interface IAiToolDispatcher
    {
        Task<ToolCallResponse> DispatchAsync(AiAgentToolDecision decision, CancellationToken cancellationToken);
    }
}
