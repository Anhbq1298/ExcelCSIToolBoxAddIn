using System;
using System.ComponentModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Threading;
using ExcelCSIToolBoxAddIn.UI.ViewModels;
using Forms = System.Windows.Forms;

namespace ExcelCSIToolBoxAddIn.UI.Views
{
    /// <summary>
    /// Code-behind for AiAgentChatTab.xaml.
    /// Hosts a WinForms textbox for more reliable keyboard input inside Excel.
    /// </summary>
    public partial class AiAgentChatTab : UserControl
    {
        private readonly Forms.TextBox _userInputTextBox;
        private AiAgentChatViewModel _viewModel;
        private bool _syncingFromViewModel;

        public AiAgentChatTab()
        {
            InitializeComponent();

            _userInputTextBox = CreateUserInputTextBox();
            UserInputHost.Child = _userInputTextBox;

            DataContextChanged += OnDataContextChanged;
            Loaded += UserControl_Loaded;
            DataContext = new AiAgentChatViewModel();
        }

        private Forms.TextBox CreateUserInputTextBox()
        {
            var textBox = new Forms.TextBox
            {
                Multiline = true,
                AcceptsReturn = true,
                ScrollBars = Forms.ScrollBars.Vertical,
                BorderStyle = Forms.BorderStyle.None,
                Font = new System.Drawing.Font("Segoe UI", 9F),
                Dock = Forms.DockStyle.Fill,
                Margin = new Forms.Padding(0)
            };

            textBox.TextChanged += UserInputTextBox_TextChanged;
            textBox.MouseDown += UserInputTextBox_MouseDown;
            textBox.KeyDown += UserInputTextBox_KeyDown;
            return textBox;
        }

        private void OnDataContextChanged(object sender, DependencyPropertyChangedEventArgs e)
        {
            if (_viewModel != null)
            {
                _viewModel.PropertyChanged -= OnViewModelPropertyChanged;
            }

            _viewModel = e.NewValue as AiAgentChatViewModel;
            if (_viewModel != null)
            {
                _viewModel.PropertyChanged += OnViewModelPropertyChanged;
                SyncTextBoxFromViewModel(_viewModel.UserInput);
            }
        }

        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            FocusInputBox();
        }

        private void UserInputTextBox_MouseDown(object sender, Forms.MouseEventArgs e)
        {
            FocusInputBox();
        }

        private void UserInputTextBox_KeyDown(object sender, Forms.KeyEventArgs e)
        {
            if (e.KeyCode == Forms.Keys.Enter && e.Control && _viewModel?.SendCommand != null)
            {
                if (_viewModel.SendCommand.CanExecute(null))
                {
                    _viewModel.SendCommand.Execute(null);
                    e.SuppressKeyPress = true;
                }
            }
        }

        private void UserInputTextBox_TextChanged(object sender, EventArgs e)
        {
            if (_syncingFromViewModel || _viewModel == null)
            {
                return;
            }

            _viewModel.UserInput = _userInputTextBox.Text;
        }

        private void OnViewModelPropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == nameof(AiAgentChatViewModel.UserInput))
            {
                SyncTextBoxFromViewModel(_viewModel?.UserInput ?? string.Empty);
            }
        }

        private void SyncTextBoxFromViewModel(string text)
        {
            string safeText = text ?? string.Empty;
            if (string.Equals(_userInputTextBox.Text, safeText, StringComparison.Ordinal))
            {
                return;
            }

            _syncingFromViewModel = true;
            _userInputTextBox.Text = safeText;
            _userInputTextBox.SelectionStart = _userInputTextBox.TextLength;
            _syncingFromViewModel = false;
        }

        private void FocusInputBox()
        {
            Window hostWindow = Window.GetWindow(this);
            if (hostWindow != null)
            {
                hostWindow.Activate();
                hostWindow.Focus();
            }

            Dispatcher.BeginInvoke(new Action(() =>
            {
                _userInputTextBox.Focus();
                _userInputTextBox.SelectionStart = _userInputTextBox.TextLength;
            }), DispatcherPriority.Input);
        }
    }
}
