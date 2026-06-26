using System;
using System.Windows;
using ExcelCSIToolBoxAddIn.UI.ViewModels;

namespace ExcelCSIToolBoxAddIn.UI.Views
{
    public partial class GetBaseReactionsWindow : Window
    {
        private bool _hasRestoredSelections;

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

            viewModel.UpdateSelectedOutputCases(LoadCasesGrid.SelectedItems, LoadCombinationsGrid.SelectedItems);
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

        protected override void OnContentRendered(EventArgs e)
        {
            base.OnContentRendered(e);
            if (_hasRestoredSelections)
            {
                return;
            }

            var viewModel = DataContext as GetBaseReactionsViewModel;
            if (viewModel != null)
            {
                viewModel.RestoreSavedSelections(LoadCasesGrid.SelectedItems, LoadCombinationsGrid.SelectedItems);
                _hasRestoredSelections = true;
            }
        }
    }
}
