using System;
using System.ComponentModel;
using System.Windows;
using System.Windows.Controls;
using ExcelCSIToolBox.Data.DTOs.CSI;
using ExcelCSIToolBoxAddIn.UI.Helpers;
using ExcelCSIToolBoxAddIn.UI.ViewModels;

namespace ExcelCSIToolBoxAddIn.UI.Views
{
    public partial class EtabsToolboxControl : UserControl
    {
        public EtabsToolboxControl()
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
            SectionShapeRenderer.Render(EtabsSectionPreviewCanvas, detail);
            EtabsSectionNameLabel.Text = detail?.Name ?? "-";
            EtabsSectionTypeLabel.Text = detail != null ? detail.ShapeType.ToString() : "";
        }

        private void HelpButton_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show(
                "This tool was developed by Mark Bui Quang Anh.",
                "About CSI Toolbox",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
        }

        private void WorkspaceNavigation_Click(object sender, RoutedEventArgs e)
        {
            var element = sender as FrameworkElement;
            NavigateToWorkspace(element == null ? null : element.Tag);
        }

        private void ProjectManagerTree_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            var item = e.NewValue as TreeViewItem;
            NavigateToWorkspace(item == null ? null : item.Tag);
        }

        private void NavigateToWorkspace(object tag)
        {
            int tabIndex;
            if (!int.TryParse(Convert.ToString(tag), out tabIndex) ||
                tabIndex < 0 ||
                tabIndex >= WorkspaceTabs.Items.Count)
            {
                return;
            }

            WorkspaceTabs.SelectedIndex = tabIndex;
        }
    }
}
