namespace ExcelCSIToolBoxAddIn.UI.ViewModels
{
    /// <summary>
    /// Represents a single chat message in the AI Agent chat history.
    /// </summary>
    public class AiAgentChatMessageViewModel : ViewModelBase
    {
        private string _role;        // "User" or "Assistant"
        private string _content;
        private bool _isTemporary;

        /// <summary>"User" or "Assistant".</summary>
        public string Role
        {
            get { return _role; }
            set
            {
                _role = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(IsUser));
            }
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

        /// <summary>True for transient UI-only messages such as the thinking indicator.</summary>
        public bool IsTemporary
        {
            get { return _isTemporary; }
            set { _isTemporary = value; OnPropertyChanged(); }
        }
    }
}
