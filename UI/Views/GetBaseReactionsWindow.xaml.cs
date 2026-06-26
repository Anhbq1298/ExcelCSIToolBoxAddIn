using System;
using System.Windows;
using ExcelCSIToolBoxAddIn.UI.ViewModels;

namespace ExcelCSIToolBoxAddIn.UI.Views
{
    public partial class GetBaseReactionsWindow : Window
    {
        public GetBaseReactionsWindow(GetBaseReactionsViewModel viewModel)
        {
            InitializeComponent();
            DataContext = viewModel ?? throw new ArgumentNullException(nameof(viewModel));
            viewModel.RequestClose += ViewModel_RequestClose;
        }

        protected override void OnClosed(EventArgs e)
        {
            if (DataContext is GetBaseReactionsViewModel viewModel)
            {
                viewModel.RequestClose -= ViewModel_RequestClose;
            }

            base.OnClosed(e);
        }

        private void ViewModel_RequestClose(object sender, EventArgs e)
        {
            Close();
        }

        private void RunButton_Click(object sender, RoutedEventArgs e)
        {
            var viewModel = DataContext as GetBaseReactionsViewModel;
            if (viewModel == null)
            {
                return;
            }

            viewModel.Run(LoadCasesGrid.SelectedItems, LoadCombinationsGrid.SelectedItems);
        }

        private void OutputCaseGrid_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            var viewModel = DataContext as GetBaseReactionsViewModel;
            if (viewModel == null)
            {
                return;
            }

            viewModel.UpdateSelectionCounts(
                LoadCasesGrid.SelectedItems.Count,
                LoadCombinationsGrid.SelectedItems.Count);
        }

        protected override void OnActivated(EventArgs e)
        {
            base.OnActivated(e);

            var viewModel = DataContext as GetBaseReactionsViewModel;
            if (viewModel != null)
            {
                viewModel.RefreshAnchorDisplay();
            }
        }
    }
}
