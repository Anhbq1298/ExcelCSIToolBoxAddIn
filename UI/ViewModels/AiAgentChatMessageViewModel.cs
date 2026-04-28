namespace ExcelCSIToolBoxAddIn.UI.ViewModels
{
    /// <summary>
    /// Represents a single chat message in the AI Agent chat history.
    /// </summary>
    public class AiAgentChatMessageViewModel : ViewModelBase
    {
        private string _role;        // "User" or "Assistant"
        private string _content;

        /// <summary>"User" or "Assistant".</summary>
        public string Role
        {
            get { return _role; }
            set { _role = value; OnPropertyChanged(); }
        }

        /// <summary>Text content of the message.</summary>
        public string Content
        {
            get { return _content; }
            set { _content = value; OnPropertyChanged(); }
        }

        /// <summary>True for user messages (used for UI alignment).</summary>
        public bool IsUser
        {
            get { return string.Equals(_role, "User", System.StringComparison.OrdinalIgnoreCase); }
        }
    }
}
