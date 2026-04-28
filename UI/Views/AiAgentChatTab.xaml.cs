using System.Windows.Controls;
using System.Windows;
using System.Windows.Threading;
using ExcelCSIToolBoxAddIn.UI.ViewModels;

namespace ExcelCSIToolBoxAddIn.UI.Views
{
    /// <summary>
    /// Code-behind for AiAgentChatTab.xaml.
    /// Sets the DataContext to a new AiAgentChatViewModel.
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
            FocusInputBox();
        }

        private void UserInputBox_PreviewMouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            FocusInputBox();
        }

        private void UserInputBox_GotKeyboardFocus(object sender, System.Windows.Input.KeyboardFocusChangedEventArgs e)
        {
            FocusInputBox();
        }

        private void FocusInputBox()
        {
            var hostWindow = Window.GetWindow(this);
            if (hostWindow != null)
            {
                hostWindow.Activate();
                hostWindow.Focus();
            }

            Dispatcher.BeginInvoke(new System.Action(() =>
            {
                if (UserInputBox == null)
                {
                    return;
                }

                UserInputBox.Focus();
                System.Windows.Input.Keyboard.Focus(UserInputBox);
                UserInputBox.CaretIndex = UserInputBox.Text?.Length ?? 0;
            }), DispatcherPriority.Input);
        }
    }
}
