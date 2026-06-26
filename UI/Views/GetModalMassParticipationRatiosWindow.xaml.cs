using System;
using System.Windows;
using ExcelCSIToolBoxAddIn.UI.ViewModels;

namespace ExcelCSIToolBoxAddIn.UI.Views
{
    public partial class GetModalMassParticipationRatiosWindow : Window
    {
        private bool _hasRestoredSelections;

        public GetModalMassParticipationRatiosWindow(GetModalMassParticipationRatiosViewModel viewModel)
        {
            InitializeComponent();
            DataContext = viewModel ?? throw new ArgumentNullException(nameof(viewModel));
            viewModel.RequestClose += ViewModel_RequestClose;
        }

        protected override void OnClosed(EventArgs e)
        {
            if (DataContext is GetModalMassParticipationRatiosViewModel viewModel)
            {
                viewModel.RequestClose -= ViewModel_RequestClose;
            }

            base.OnClosed(e);
        }

        protected override void OnActivated(EventArgs e)
        {
            base.OnActivated(e);

            var viewModel = DataContext as GetModalMassParticipationRatiosViewModel;
            if (viewModel != null)
            {
                viewModel.RefreshAnchorDisplay();
            }
        }

        private void ViewModel_RequestClose(object sender, EventArgs e)
        {
            Close();
        }

        private void ModalLoadCasesGrid_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            var viewModel = DataContext as GetModalMassParticipationRatiosViewModel;
            if (viewModel != null)
            {
                viewModel.UpdateSelectedOutputCases(ModalLoadCasesGrid.SelectedItems);
            }
        }

        private void RunButton_Click(object sender, RoutedEventArgs e)
        {
            var viewModel = DataContext as GetModalMassParticipationRatiosViewModel;
            if (viewModel != null)
            {
                viewModel.Run(ModalLoadCasesGrid.SelectedItems);
            }
        }

        protected override void OnContentRendered(EventArgs e)
        {
            base.OnContentRendered(e);
            if (_hasRestoredSelections)
            {
                return;
            }

            var viewModel = DataContext as GetModalMassParticipationRatiosViewModel;
            if (viewModel != null)
            {
                viewModel.RestoreSavedSelections(ModalLoadCasesGrid.SelectedItems);
                _hasRestoredSelections = true;
            }
        }
    }
}
