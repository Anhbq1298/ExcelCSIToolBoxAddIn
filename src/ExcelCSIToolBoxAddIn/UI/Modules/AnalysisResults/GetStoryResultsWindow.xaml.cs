using System;
using System.Windows;
using ExcelCSIToolBoxAddIn.AddIn;
using ExcelCSIToolBoxAddIn.UI.ViewModels;

namespace ExcelCSIToolBoxAddIn.UI.Views
{
    public partial class GetStoryResultsWindow : Window
    {
        private bool _hasRestoredSelections;

        public GetStoryResultsWindow(GetStoryResultsViewModel viewModel)
        {
            InitializeComponent();
            DataContext = viewModel ?? throw new ArgumentNullException(nameof(viewModel));
            viewModel.RequestClose += ViewModel_RequestClose;
            viewModel.RequestHide += ViewModel_RequestHide;
            viewModel.RequestShow += ViewModel_RequestShow;
        }

        protected override void OnClosed(EventArgs e)
        {
            if (DataContext is GetStoryResultsViewModel viewModel)
            {
                viewModel.RequestClose -= ViewModel_RequestClose;
                viewModel.RequestHide -= ViewModel_RequestHide;
                viewModel.RequestShow -= ViewModel_RequestShow;
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

        private void ViewModel_RequestHide(object sender, EventArgs e)
        {
            this.Hide();
        }

        private void ViewModel_RequestShow(object sender, EventArgs e)
        {
            ModelessWpfWindowService.Show(this);
        }

        private void OutputCaseGrid_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            var viewModel = DataContext as GetStoryResultsViewModel;
            if (viewModel != null)
            {
                viewModel.UpdateSelectedOutputCases(LoadCasesGrid.SelectedItems, LoadCombinationsGrid.SelectedItems);
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

        protected override void OnContentRendered(EventArgs e)
        {
            base.OnContentRendered(e);
            if (_hasRestoredSelections)
            {
                return;
            }

            var viewModel = DataContext as GetStoryResultsViewModel;
            if (viewModel != null)
            {
                viewModel.RestoreSavedSelections(LoadCasesGrid.SelectedItems, LoadCombinationsGrid.SelectedItems);
                _hasRestoredSelections = true;
            }
        }
    }
}
