using System;
using System.Collections.ObjectModel;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using ExcelCSIToolBox.AI.Agent;
using ExcelCSIToolBox.AI.Mcp.Client;
using ExcelCSIToolBox.AI.Mcp.Server;
using ExcelCSIToolBox.AI.Ollama;
using ExcelCSIToolBox.Core.Abstractions.CSI;
using ExcelCSIToolBox.Infrastructure.CSISapModel.ReadOnly;
using ExcelCSIToolBox.Infrastructure.Etabs;
using ExcelCSIToolBox.Infrastructure.Sap2000;

namespace ExcelCSIToolBoxAddIn.UI.ViewModels
{
    /// <summary>
    /// ViewModel for the AI Agent Chat tab.
    ///
    /// Wires up:
    ///   - OllamaChatService
    ///   - CsiReadOnly* Infrastructure services
    ///   - LocalMcpServer → LocalMcpClient
    ///   - AiAgentOrchestrator
    ///
    /// SendCommand runs the full agent loop asynchronously and posts results
    /// back to the WPF UI thread.
    /// </summary>
    public class AiAgentChatViewModel : ViewModelBase
    {
        // ── Backing fields ────────────────────────────────────────────────────────

        private string                       _userInput  = string.Empty;
        private string                       _statusText = "Ready";
        private string                       _currentModelName = OllamaChatService.DefaultModel;
        private string                       _sap2000ConnectionStatus = "Not attached";
        private string                       _etabsConnectionStatus = "Not attached";
        private bool                         _isBusy     = false;
        private AiAgentToolTraceViewModel    _lastToolTrace;

        // ── Agent infrastructure ──────────────────────────────────────────────────

        private readonly AiAgentOrchestrator _orchestrator;
        private readonly ICSISapModelConnectionService _etabsConnectionService;
        private readonly ICSISapModelConnectionService _sap2000ConnectionService;

        // ── Constructor ───────────────────────────────────────────────────────────

        public AiAgentChatViewModel()
            : this(new EtabsConnectionService(), new Sap2000ConnectionService())
        {
        }

        public AiAgentChatViewModel(
            ICSISapModelConnectionService etabsConnectionService,
            ICSISapModelConnectionService sap2000ConnectionService)
        {
            _etabsConnectionService = etabsConnectionService ?? throw new ArgumentNullException(nameof(etabsConnectionService));
            _sap2000ConnectionService = sap2000ConnectionService ?? throw new ArgumentNullException(nameof(sap2000ConnectionService));

            // Build read-only Infrastructure services.
            ICsiReadOnlyConnectionService connectionService = new CsiReadOnlyConnectionService();
            ICsiReadOnlySelectionService  selectionService  = new CsiReadOnlySelectionService();
            ICsiReadOnlyFrameService      frameService      = new CsiReadOnlyFrameService();

            // Build local MCP server and client.
            LocalMcpServer mcpServer = new LocalMcpServer(connectionService, selectionService, frameService);
            LocalMcpClient mcpClient = new LocalMcpClient(mcpServer);

            // Build Ollama chat service.
            OllamaChatService ollamaService = new OllamaChatService();

            // Build orchestrator.
            _orchestrator = new AiAgentOrchestrator(ollamaService, mcpClient);

            // Initialise observable state.
            Messages       = new ObservableCollection<AiAgentChatMessageViewModel>();
            _lastToolTrace = new AiAgentToolTraceViewModel();

            // Commands.
            SendCommand  = new AiRelayCommand(ExecuteSend,  CanExecuteSend);
            ClearCommand = new AiRelayCommand(ExecuteClear);

            RefreshConnectionStatuses();
        }

        // ── Observable properties ─────────────────────────────────────────────────

        public ObservableCollection<AiAgentChatMessageViewModel> Messages { get; }

        public string UserInput
        {
            get { return _userInput; }
            set
            {
                _userInput = value;
                OnPropertyChanged();
                ((AiRelayCommand)SendCommand).RaiseCanExecuteChanged();
            }
        }

        public string StatusText
        {
            get { return _statusText; }
            private set { _statusText = value; OnPropertyChanged(); }
        }

        public string CurrentModelName
        {
            get { return _currentModelName; }
            set { _currentModelName = string.IsNullOrWhiteSpace(value) ? "Not selected" : value; OnPropertyChanged(); }
        }

        public string Sap2000ConnectionStatus
        {
            get { return _sap2000ConnectionStatus; }
            set { _sap2000ConnectionStatus = string.IsNullOrWhiteSpace(value) ? "Not attached" : value; OnPropertyChanged(); }
        }

        public string EtabsConnectionStatus
        {
            get { return _etabsConnectionStatus; }
            set { _etabsConnectionStatus = string.IsNullOrWhiteSpace(value) ? "Not attached" : value; OnPropertyChanged(); }
        }

        public bool IsBusy
        {
            get { return _isBusy; }
            private set
            {
                _isBusy = value;
                OnPropertyChanged();
                ((AiRelayCommand)SendCommand).RaiseCanExecuteChanged();
            }
        }

        public AiAgentToolTraceViewModel LastToolTrace
        {
            get { return _lastToolTrace; }
            private set { _lastToolTrace = value; OnPropertyChanged(); }
        }

        // ── Commands ──────────────────────────────────────────────────────────────

        public ICommand SendCommand  { get; }
        public ICommand ClearCommand { get; }

        // ── Command implementations ───────────────────────────────────────────────

        private bool CanExecuteSend(object _)
        {
            return !_isBusy && !string.IsNullOrWhiteSpace(_userInput);
        }

        private async void ExecuteSend(object _)
        {
            string userMessage = _userInput.Trim();
            if (string.IsNullOrWhiteSpace(userMessage)) return;

            RefreshConnectionStatuses();

            // Add the user message to the chat history.
            Messages.Add(new AiAgentChatMessageViewModel { Role = "User", Content = userMessage });
            var thinkingMessage = new AiAgentChatMessageViewModel
            {
                Role = "Assistant",
                Content = "...",
                IsTemporary = true
            };
            Messages.Add(thinkingMessage);
            UserInput  = string.Empty;
            IsBusy     = true;
            StatusText = "Thinking…";
            _lastToolTrace.Clear();

            try
            {
                AiAgentResponse response = await _orchestrator.SendAsync(
                    userMessage,
                    CancellationToken.None);

                Messages.Remove(thinkingMessage);

                // Add assistant reply.
                Messages.Add(new AiAgentChatMessageViewModel
                {
                    Role    = "Assistant",
                    Content = response.AssistantText
                });

                // Update tool trace panel.
                if (response.ToolWasCalled && response.ToolResponse != null)
                {
                    _lastToolTrace.ToolWasCalled  = true;
                    _lastToolTrace.ToolName        = response.ToolName;
                    _lastToolTrace.ToolArguments   = response.ToolArgumentsJson;
                    _lastToolTrace.ToolSucceeded   = response.ToolResponse.Success;
                    _lastToolTrace.ToolMessage     = response.ToolResponse.Message;
                    _lastToolTrace.ToolResultJson  = response.ToolResponse.ResultJson ?? "(none)";
                }

                StatusText = "Ready";
            }
            catch (Exception ex)
            {
                Messages.Remove(thinkingMessage);

                Messages.Add(new AiAgentChatMessageViewModel
                {
                    Role    = "Assistant",
                    Content = "⚠️ Error: " + ex.Message +
                              "\n\nMake sure Ollama is running and the model is pulled."
                });
                StatusText = "Error";
            }
            finally
            {
                IsBusy = false;
            }
        }

        private void ExecuteClear(object _)
        {
            Messages.Clear();
            _lastToolTrace.Clear();
            StatusText = "Ready";
        }

        private void RefreshConnectionStatuses()
        {
            EtabsConnectionStatus = IsProductAttached(_etabsConnectionService)
                ? "Attached"
                : "Not attached";

            Sap2000ConnectionStatus = IsProductAttached(_sap2000ConnectionService)
                ? "Attached"
                : "Not attached";
        }

        private static bool IsProductAttached(ICSISapModelConnectionService connectionService)
        {
            try
            {
                var result = connectionService.TryAttachToRunningInstance();
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

    // ── Minimal AiRelayCommand ────────────────────────────────────────────────────

    /// <summary>
    /// Simple relay command for the AI Agent tab.
    /// Named AiRelayCommand to avoid collision with the project-wide RelayCommand in Core.
    /// </summary>
    internal sealed class AiRelayCommand : ICommand
    {
        private readonly Action<object>     _execute;
        private readonly Func<object, bool> _canExecute;

        public AiRelayCommand(Action<object> execute, Func<object, bool> canExecute = null)
        {
            _execute    = execute    ?? throw new ArgumentNullException(nameof(execute));
            _canExecute = canExecute;
        }

        public event EventHandler CanExecuteChanged;

        public bool CanExecute(object parameter)
        {
            return _canExecute == null || _canExecute(parameter);
        }

        public void Execute(object parameter)
        {
            _execute(parameter);
        }

        public void RaiseCanExecuteChanged()
        {
            CanExecuteChanged?.Invoke(this, EventArgs.Empty);
        }
    }
}
