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
            PreviewKeyDown += OnPreviewKeyDown;
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
            PreviewKeyDown -= OnPreviewKeyDown;
        }

        private void OnPreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            IInputElement input = FindTextInputAncestor(e.OriginalSource as DependencyObject);
            if (input != null && !HasKeyboardFocusWithin(input))
            {
                Keyboard.Focus(input);
            }
        }

        private void OnPreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (IsTextInputFocused())
            {
                return;
            }
        }

        private static bool IsTextInputFocused()
        {
            IInputElement focusedElement = Keyboard.FocusedElement;
            return focusedElement is TextBox ||
                   focusedElement is PasswordBox ||
                   focusedElement is ComboBox ||
                   focusedElement is DatePicker;
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

        private static bool HasKeyboardFocusWithin(IInputElement input)
        {
            UIElement uiElement = input as UIElement;
            if (uiElement != null)
            {
                return uiElement.IsKeyboardFocusWithin;
            }

            ContentElement contentElement = input as ContentElement;
            if (contentElement != null)
            {
                return contentElement.IsKeyboardFocusWithin;
            }

            return input.IsKeyboardFocused;
        }
    }
}
