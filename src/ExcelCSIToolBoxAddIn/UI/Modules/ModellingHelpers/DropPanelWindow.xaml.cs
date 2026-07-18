using System;
using System.ComponentModel;
using System.Windows;
using System.Windows.Input;
using System.Windows.Threading;
using ExcelCSIToolBoxAddIn.UI.ViewModels;

namespace ExcelCSIToolBoxAddIn.UI.Views
{
    public partial class DropPanelWindow : Window
    {
        public DropPanelWindow(DropPanelViewModel viewModel)
        {
            InitializeComponent();
            DataContext = viewModel ?? throw new ArgumentNullException(nameof(viewModel));
            Activated += OnActivated;
            Closing += OnClosing;
            Closed += OnClosed;
        }

        private void OnActivated(object sender, EventArgs e)
        {
            DropPanelViewModel viewModel = DataContext as DropPanelViewModel;
            if (viewModel != null)
            {
                viewModel.RefreshModelInputs(false);
            }

            if (!IsKeyboardFocusWithin)
            {
                Dispatcher.BeginInvoke(
                    new Action(delegate
                    {
                        DropThicknessTextBox.Focus();
                        Keyboard.Focus(DropThicknessTextBox);
                    }),
                    DispatcherPriority.Input);
            }
        }

        private void OnClosing(object sender, CancelEventArgs e)
        {
            DropPanelViewModel viewModel = DataContext as DropPanelViewModel;
            if (viewModel != null && !viewModel.TryCloseWindow())
            {
                e.Cancel = true;
            }
        }

        private void OnClosed(object sender, EventArgs e)
        {
            Activated -= OnActivated;
            Closing -= OnClosing;
            Closed -= OnClosed;
        }
    }
}
