using System;
using System.Windows;
using ExcelCSIToolBoxAddIn.UI.ViewModels;

namespace ExcelCSIToolBoxAddIn.UI.Views
{
    public partial class OutputTableExportOptionsWindow : Window
    {
        private bool _hasRestoredSelections;

        public OutputTableExportOptionsWindow(OutputTableExportOptionsViewModel viewModel)
        {
            InitializeComponent();
            DataContext = viewModel ?? throw new ArgumentNullException(nameof(viewModel));
            LoadCasesGrid.SelectionMode = viewModel.AllowMultipleCases
                ? System.Windows.Controls.DataGridSelectionMode.Extended
                : System.Windows.Controls.DataGridSelectionMode.Single;
            LoadCombinationsGrid.SelectionMode = viewModel.AllowMultipleCases
                ? System.Windows.Controls.DataGridSelectionMode.Extended
                : System.Windows.Controls.DataGridSelectionMode.Single;
            viewModel.RequestClose += ViewModel_RequestClose;
        }

        protected override void OnClosed(EventArgs e)
        {
            if (DataContext is OutputTableExportOptionsViewModel viewModel)
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
            var viewModel = DataContext as OutputTableExportOptionsViewModel;
            if (viewModel == null)
            {
                return;
            }

            viewModel.Run(LoadCasesGrid.SelectedItems, LoadCombinationsGrid.SelectedItems);
        }

        private void OutputCaseGrid_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            var viewModel = DataContext as OutputTableExportOptionsViewModel;
            if (viewModel == null)
            {
                return;
            }

            if (!viewModel.AllowMultipleCases)
            {
                if (sender == LoadCasesGrid && LoadCasesGrid.SelectedItems.Count > 0)
                {
                    LoadCombinationsGrid.SelectedItem = null;
                }
                else if (sender == LoadCombinationsGrid && LoadCombinationsGrid.SelectedItems.Count > 0)
                {
                    LoadCasesGrid.SelectedItem = null;
                }
            }

            viewModel.UpdateSelectedOutputCases(LoadCasesGrid.SelectedItems, LoadCombinationsGrid.SelectedItems);
        }

        protected override void OnActivated(EventArgs e)
        {
            base.OnActivated(e);

            var viewModel = DataContext as OutputTableExportOptionsViewModel;
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

            var viewModel = DataContext as OutputTableExportOptionsViewModel;
            if (viewModel != null)
            {
                viewModel.InitializeForDialog();

                if (viewModel.CaseComboSelectorVisibility == Visibility.Visible)
                {
                    // Unsubscribe selection changed events to avoid triggering updates/resets during restoration
                    LoadCasesGrid.SelectionChanged -= OutputCaseGrid_SelectionChanged;
                    LoadCombinationsGrid.SelectionChanged -= OutputCaseGrid_SelectionChanged;

                    try
                    {
                        viewModel.RestoreSavedSelections(
                            items => SetGridSelection(LoadCasesGrid, items),
                            items => SetGridSelection(LoadCombinationsGrid, items)
                        );
                    }
                    finally
                    {
                        // Re-subscribe selection changed events
                        LoadCasesGrid.SelectionChanged += OutputCaseGrid_SelectionChanged;
                        LoadCombinationsGrid.SelectionChanged += OutputCaseGrid_SelectionChanged;
                    }

                    viewModel.UpdateSelectedOutputCases(LoadCasesGrid.SelectedItems, LoadCombinationsGrid.SelectedItems);
                }

                _hasRestoredSelections = true;
            }
        }

        private void SetGridSelection(System.Windows.Controls.DataGrid grid, System.Collections.Generic.IEnumerable<BaseReactionOutputCaseViewModel> itemsToSelect)
        {
            if (grid == null) return;

            grid.SelectedItem = null;

            if (itemsToSelect == null)
            {
                return;
            }

            if (grid.SelectionMode == System.Windows.Controls.DataGridSelectionMode.Single)
            {
                foreach (var item in itemsToSelect)
                {
                    grid.SelectedItem = item;
                    break;
                }
            }
            else
            {
                foreach (var item in itemsToSelect)
                {
                    grid.SelectedItems.Add(item);
                }
            }
        }
    }
}
