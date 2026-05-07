using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using ExcelCSIToolBox.AI.Mcp.Safety;

namespace ExcelCSIToolBoxAddIn.AddIn
{
    /// <summary>
    /// WPF confirmation guard for AI-initiated model mutations.
    /// </summary>
    internal sealed class WpfMutationGuard : IMutationGuard
    {
        public Task<bool> ConfirmAsync(string toolName, string summary, CancellationToken ct)
        {
            if (ct.IsCancellationRequested)
            {
                return Task.FromResult(false);
            }

            MessageBoxResult result = MessageBox.Show(
                $"{toolName}\n\n{summary}",
                "Confirm CSI Model Change",
                MessageBoxButton.OKCancel,
                MessageBoxImage.Warning);

            return Task.FromResult(result == MessageBoxResult.OK);
        }
    }
}
