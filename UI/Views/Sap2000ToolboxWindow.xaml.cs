using System.ComponentModel;
using System.Windows;
using ExcelCSIToolBoxAddIn.Data.DTOs;
using ExcelCSIToolBoxAddIn.UI.Helpers;
using ExcelCSIToolBoxAddIn.UI.ViewModels;

namespace ExcelCSIToolBoxAddIn.UI.Views
{
    /// <summary>
    /// Interaction logic for Sap2000ToolboxWindow.xaml
    /// </summary>
    public partial class Sap2000ToolboxWindow : Window
    {
        public Sap2000ToolboxWindow()
        {
            InitializeComponent();
            RenderSectionPreview(null);
            DataContextChanged += OnDataContextChanged;
        }

        private void OnDataContextChanged(object sender, DependencyPropertyChangedEventArgs e)
        {
            if (e.OldValue is CsiToolboxViewModel oldVm)
                oldVm.PropertyChanged -= OnViewModelPropertyChanged;
            if (e.NewValue is CsiToolboxViewModel newVm)
                newVm.PropertyChanged += OnViewModelPropertyChanged;
        }

        private void OnViewModelPropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == nameof(CsiToolboxViewModel.SelectedFrameSectionDetail))
            {
                var vm = (CsiToolboxViewModel)sender;
                RenderSectionPreview(vm.SelectedFrameSectionDetail);
            }
        }

        private void RenderSectionPreview(CSISapModelFrameSectionDetailDTO detail)
        {
            SectionShapeRenderer.Render(Sap2000SectionPreviewCanvas, detail);
            Sap2000SectionNameLabel.Text = detail?.Name ?? "-";
            Sap2000SectionTypeLabel.Text = detail != null ? detail.ShapeType.ToString() : "";
        }

        private void HelpButton_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show(
                "This tool was developed by Mark Bui Quang Anh.",
                "About SAP2000 Toolbox",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
        }
    }
}
