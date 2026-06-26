using System;
using System.Windows;
using ExcelCSIToolBoxAddIn.UI.ViewModels;

namespace ExcelCSIToolBoxAddIn.UI.Views
{
    public partial class GetStoryResultsWindow : Window
    {
        public GetStoryResultsWindow(GetStoryResultsViewModel viewModel)
        {
            InitializeComponent();
            DataContext = viewModel ?? throw new ArgumentNullException(nameof(viewModel));
            viewModel.RequestClose += ViewModel_RequestClose;
        }

        protected override void OnClosed(EventArgs e)
        {
            if (DataContext is GetStoryResultsViewModel viewModel)
            {
                viewModel.RequestClose -= ViewModel_RequestClose;
            }

            base.OnClosed(e);
        }

        protected override void OnActivated(EventArgs e)
        {
            base.OnActivated(e);

            var viewModel = DataContext as GetStoryResultsViewModel;
            if (viewModel != null)
            {
                viewModel.RefreshAnchorDisplay();
            }
        }

        private void ViewModel_RequestClose(object sender, EventArgs e)
        {
            Close();
        }

        private void OutputCaseGrid_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            var viewModel = DataContext as GetStoryResultsViewModel;
            if (viewModel != null)
            {
                viewModel.UpdateSelectionCounts(
                    LoadCasesGrid.SelectedItems.Count,
                    LoadCombinationsGrid.SelectedItems.Count);
            }
        }

        private void RunButton_Click(object sender, RoutedEventArgs e)
        {
            var viewModel = DataContext as GetStoryResultsViewModel;
            if (viewModel != null)
            {
                viewModel.Run(LoadCasesGrid.SelectedItems, LoadCombinationsGrid.SelectedItems);
            }
        }
    }
}
