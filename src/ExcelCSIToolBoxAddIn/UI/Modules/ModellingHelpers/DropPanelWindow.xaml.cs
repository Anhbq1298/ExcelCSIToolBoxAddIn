using System;
using System.ComponentModel;
using System.Windows;
using ExcelCSIToolBoxAddIn.UI.ViewModels;

namespace ExcelCSIToolBoxAddIn.UI.Views
{
    public partial class DropPanelWindow : Window
    {
        public DropPanelWindow(DropPanelViewModel viewModel)
        {
            InitializeComponent();
            DataContext = viewModel ?? throw new ArgumentNullException(nameof(viewModel));
            Closing += OnClosing;
        }

        private void OnClosing(object sender, CancelEventArgs e)
        {
            DropPanelViewModel viewModel = DataContext as DropPanelViewModel;
            if (viewModel != null && !viewModel.TryCloseWindow())
            {
                e.Cancel = true;
            }
        }
    }
}
