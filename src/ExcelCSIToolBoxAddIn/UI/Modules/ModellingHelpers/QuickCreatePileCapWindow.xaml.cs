using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;
using ExcelCSIToolBoxAddIn.UI.ViewModels;

namespace ExcelCSIToolBoxAddIn.UI.Views
{
    public partial class QuickCreatePileCapWindow : Window
    {
        public QuickCreatePileCapWindow(QuickCreatePileCapViewModel viewModel)
        {
            InitializeComponent();
            DataContext = viewModel ?? throw new ArgumentNullException("viewModel");
            Loaded += OnLoaded;
            Closed += OnClosed;
            PreviewMouseDown += OnPreviewMouseDown;
        }

        private void OnLoaded(object sender, RoutedEventArgs e)
        {
            Dispatcher.BeginInvoke(new Action(delegate
            {
                Activate();
                Focus();
            }), DispatcherPriority.ApplicationIdle);
        }

        private void OnClosed(object sender, EventArgs e)
        {
            Loaded -= OnLoaded;
            Closed -= OnClosed;
            PreviewMouseDown -= OnPreviewMouseDown;
        }

        private void OnPreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            IInputElement input = FindTextInputAncestor(e.OriginalSource as DependencyObject);
            EnableManualSpacingIfNeeded(input);
        }

        private bool EnableManualSpacingIfNeeded(IInputElement input)
        {
            TextBox textBox = input as TextBox;
            if (textBox == null || textBox.IsEnabled ||
                (textBox != PileSpacingTextBox && textBox != SpacingXTextBox && textBox != SpacingYTextBox))
            {
                return false;
            }

            QuickCreatePileCapViewModel viewModel = DataContext as QuickCreatePileCapViewModel;
            if (viewModel != null && viewModel.IsAutomaticSpacing)
            {
                viewModel.IsAutomaticSpacing = false;
                Dispatcher.BeginInvoke(new Action(delegate
                {
                    FocusTextInput(textBox);
                    textBox.SelectAll();
                }), DispatcherPriority.Input);
                return true;
            }

            return false;
        }

        private static IInputElement FindTextInputAncestor(DependencyObject source)
        {
            while (source != null)
            {
                IInputElement input = source as IInputElement;
                if (input is TextBox ||
                    input is PasswordBox ||
                    input is ComboBox ||
                    input is DatePicker)
                {
                    return input;
                }

                source = VisualTreeHelper.GetParent(source);
            }

            return null;
        }

        private void FocusTextInput(IInputElement input)
        {
            Control control = input as Control;
            if (control != null)
            {
                control.Focus();
            }

            Keyboard.Focus(input);
        }
    }
}
