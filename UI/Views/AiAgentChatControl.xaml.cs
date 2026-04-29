using System.Collections.Specialized;
using System.Windows.Controls;
using System.Windows.Input;
using ExcelCSIToolBoxAddIn.UI.ViewModels;

namespace ExcelCSIToolBoxAddIn.UI.Views
{
    public partial class AiAgentChatControl : UserControl
    {
        private readonly AiAgentChatViewModel _viewModel;

        public AiAgentChatControl()
        {
            InitializeComponent();

            _viewModel = new AiAgentChatViewModel();
            DataContext = _viewModel;

            _viewModel.Messages.CollectionChanged += Messages_CollectionChanged;
        }

        private void Messages_CollectionChanged(object sender, NotifyCollectionChangedEventArgs e)
        {
            Dispatcher.BeginInvoke(new System.Action(ScrollConversationToEnd));
        }

        private void InputTextBox_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key != Key.Enter || Keyboard.Modifiers != ModifierKeys.Control)
            {
                return;
            }

            if (_viewModel.SendCommand.CanExecute(null))
            {
                _viewModel.SendCommand.Execute(null);
                e.Handled = true;
            }
        }

        private void ScrollConversationToEnd()
        {
            ConversationScrollViewer.ScrollToEnd();
        }
    }
}
