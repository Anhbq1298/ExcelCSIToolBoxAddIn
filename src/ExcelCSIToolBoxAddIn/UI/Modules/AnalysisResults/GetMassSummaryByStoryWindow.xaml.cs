using System;
using System.Windows;
using ExcelCSIToolBoxAddIn.AddIn;
using ExcelCSIToolBoxAddIn.UI.ViewModels;

namespace ExcelCSIToolBoxAddIn.UI.Views
{
    public partial class GetMassSummaryByStoryWindow : Window
    {
        public GetMassSummaryByStoryWindow(GetMassSummaryByStoryViewModel viewModel)
        {
            InitializeComponent();
            DataContext = viewModel ?? throw new ArgumentNullException(nameof(viewModel));
            viewModel.RequestClose += ViewModel_RequestClose;
            viewModel.RequestHide += ViewModel_RequestHide;
            viewModel.RequestShow += ViewModel_RequestShow;
        }

        protected override void OnClosed(EventArgs e)
        {
            var viewModel = DataContext as GetMassSummaryByStoryViewModel;
            if (viewModel != null)
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
            var viewModel = DataContext as GetMassSummaryByStoryViewModel;
            if (viewModel != null) viewModel.RefreshAnchorDisplay();
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
    }
}
