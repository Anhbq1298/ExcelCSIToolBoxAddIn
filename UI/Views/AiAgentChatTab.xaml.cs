using System.Windows.Controls;
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

        private void UserInputBox_PreviewMouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            // Explicitly focus and capture the keyboard to help Excel interop
            var textBox = sender as TextBox;
            if (textBox != null)
            {
                textBox.Focus();
                System.Windows.Input.Keyboard.Focus(textBox);
            }
        }
    }
}
