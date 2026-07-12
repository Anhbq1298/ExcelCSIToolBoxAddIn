using System.Collections.Specialized;
using System.Windows.Controls;
using System.Windows.Input;
using System;
using ExcelCSIToolBox.Core.Abstractions;
using ExcelCSIToolBoxAddIn.UI.ViewModels;

namespace ExcelCSIToolBoxAddIn.UI.Views
{
    public partial class AiAgentChatControl : UserControl
    {
        private readonly AiAgentChatViewModel _viewModel;
        private readonly IThreadDispatcher _threadDispatcher;

        public AiAgentChatControl(AiAgentChatViewModel viewModel, IThreadDispatcher threadDispatcher)
        {
            InitializeComponent();

            _viewModel = viewModel ?? throw new ArgumentNullException(nameof(viewModel));
            _threadDispatcher = threadDispatcher ?? throw new ArgumentNullException(nameof(threadDispatcher));
            DataContext = _viewModel;

            _viewModel.Messages.CollectionChanged += Messages_CollectionChanged;
        }

        private void Messages_CollectionChanged(object sender, NotifyCollectionChangedEventArgs e)
        {
            _threadDispatcher.InvokeOnUiThread(ScrollConversationToEnd);
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
