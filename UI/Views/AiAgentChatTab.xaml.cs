using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Threading;
using ExcelCSIToolBoxAddIn.UI.ViewModels;

namespace ExcelCSIToolBoxAddIn.UI.Views
{
    /// <summary>
    /// Code-behind for AiAgentChatTab.xaml.
    /// </summary>
    public partial class AiAgentChatTab : UserControl
    {
        public AiAgentChatTab()
        {
            InitializeComponent();
            DataContext = new AiAgentChatViewModel();
        }

        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            Dispatcher.BeginInvoke(new Action(() =>
            {
                UserInputBox.Focus();
                Keyboard.Focus(UserInputBox);
            }), DispatcherPriority.Input);
        }

        private void UserInputBox_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter && Keyboard.Modifiers.HasFlag(ModifierKeys.Control))
            {
                var viewModel = DataContext as AiAgentChatViewModel;
                if (viewModel?.SendCommand != null && viewModel.SendCommand.CanExecute(null))
                {
                    viewModel.SendCommand.Execute(null);
                    e.Handled = true;
                }
            }
        }
    }
}
