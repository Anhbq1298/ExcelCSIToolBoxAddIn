namespace ExcelCSIToolBoxAddIn.UI.ViewModels
{
    /// <summary>
    /// Holds the tool trace from the last tool call for the debug panel in the UI.
    /// </summary>
    public class AiAgentToolTraceViewModel : ViewModelBase
    {
        private bool   _toolWasCalled;
        private string _toolName;
        private string _toolArguments;
        private bool   _toolSucceeded;
        private string _toolMessage;
        private string _toolResultJson;
        private string _routingReason;

        /// <summary>Whether the last response involved a tool call.</summary>
        public bool ToolWasCalled
        {
            get { return _toolWasCalled; }
            set { _toolWasCalled = value; OnPropertyChanged(); }
        }

        /// <summary>Name of the tool that was called.</summary>
        public string ToolName
        {
            get { return _toolName; }
            set { _toolName = value; OnPropertyChanged(); }
        }

        /// <summary>JSON arguments that were sent to the tool.</summary>
        public string ToolArguments
        {
            get { return _toolArguments; }
            set { _toolArguments = value; OnPropertyChanged(); }
        }

        /// <summary>Whether the tool call succeeded.</summary>
        public bool ToolSucceeded
        {
            get { return _toolSucceeded; }
            set { _toolSucceeded = value; OnPropertyChanged(); }
        }

        /// <summary>Tool result message (success or error text).</summary>
        public string ToolMessage
        {
            get { return _toolMessage; }
            set { _toolMessage = value; OnPropertyChanged(); }
        }

        /// <summary>Raw JSON returned by the tool.</summary>
        public string ToolResultJson
        {
            get { return _toolResultJson; }
            set { _toolResultJson = value; OnPropertyChanged(); }
        }

        /// <summary>Short explanation of how the agent routed the request.</summary>
        public string RoutingReason
        {
            get { return _routingReason; }
            set { _routingReason = value; OnPropertyChanged(); }
        }

        /// <summary>Reset all fields to empty state.</summary>
        public void Clear()
        {
            ToolWasCalled   = false;
            ToolName        = string.Empty;
            ToolArguments   = string.Empty;
            ToolSucceeded   = false;
            ToolMessage     = string.Empty;
            ToolResultJson  = string.Empty;
            RoutingReason   = string.Empty;
        }
    }
}
