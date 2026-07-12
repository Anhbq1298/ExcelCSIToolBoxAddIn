using System;
using System.ComponentModel;
using System.Windows;
using System.Windows.Controls;
using ExcelCSIToolBox.Core.Contracts.CSI;
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
            {
                newVm.PropertyChanged += OnViewModelPropertyChanged;
            }
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

        private void ProjectManagerTree_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            var item = e.NewValue as TreeViewItem;
            var viewModel = DataContext as CsiToolboxViewModel;
            var pageIndex = item == null ? null : item.Tag as string;
            if (viewModel != null && !string.IsNullOrWhiteSpace(pageIndex) &&
                viewModel.SelectWorkspacePageCommand.CanExecute(pageIndex))
            {
                viewModel.SelectWorkspacePageCommand.Execute(pageIndex);
            }
        }

    }

    public class TableHeaderToolTipConverter : System.Windows.Data.IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            if (value == null)
                return null;

            string tableName = value.ToString().Trim();

            switch (tableName)
            {
                case "Base Reactions":
                    return "Format: Output Case | Case Type | Step Type | Step Number | FX | FY | FZ | MX | MY | MZ | X | Y | Z";
                case "Joint Displacements":
                case "Joint Displacements - Absolute":
                    return "Format: Unique Name | Output Case | Step Type | Step Number | U1 | U2 | U3 | R1 | R2 | R3";
                case "Joint Drifts":
                    return "Format: Story | Label | Unique Name | Output Case | Step Type | Step Number | Displacement X | Displacement Y | Drift X | Drift Y";
                case "Joint Reactions":
                case "Joint Design Reactions":
                    return "Format: Unique Name | Output Case | Step Type | Step Number | FX | FY | FZ | MX | MY | MZ";
                case "Joint Velocities - Relative":
                case "Joint Velocities - Absolute":
                    return "Format: Unique Name | Output Case | Step Type | Step Number | U1 | U2 | U3 | R1 | R2 | R3";
                case "Joint Accelerations - Relative":
                case "Joint Accelerations - Absolute":
                    return "Format: Unique Name | Output Case | Step Type | Step Number | U1 | U2 | U3 | R1 | R2 | R3";
                case "Assembled Joint Masses":
                    return "Format: Unique Name | U1 | U2 | U3 | R1 | R2 | R3";
                case "Story Forces":
                    return "Format: Story | Output Case | Location | P | V2 | V3 | T | M2 | M3";
                case "Diaphragm Forces":
                    return "Format: Story | Diaphragm | Output Case | Location | P | Vx | Vy | T | Mx | My";
                case "Story Stiffness":
                    return "Format: Story | Output Case | Direction | Stiffness X | Stiffness Y";
                case "Shear Gravity Ratios":
                case "Stiffness Gravity Ratios":
                    return "Format: Story | Ratio";
                case "Centers Of Mass And Rigidity":
                    return "Format: Story | Diaphragm | Mass X | Mass Y | COM X | COM Y | COR X | COR Y";
                case "Element Forces - Columns":
                case "Element Forces - Beams":
                case "Element Forces - Braces":
                    return "Format: Frame | Station | Output Case | Case Type | Step Type | Step Number | P | V2 | V3 | T | M2 | M3";
                case "Element Joint Forces - Frame":
                    return "Format: Frame | Joint | Output Case | Case Type | Step Type | Step Number | ElementType | FX | FY | FZ | MX | MY | MZ";
                case "Element Forces - Area Shells":
                    return "Format: Area | Joint | Output Case | Case Type | Step Type | Step Number | F11 | F22 | F12 | FMax | FMin | FAngle | V13 | V23 | VMax | VAngle | M11 | M22 | M12 | MMax | MMin | MAngle";
                case "Element Stresses - Area Shells":
                    return "Format: Area | Joint | Output Case | Case Type | Step Type | Step Number | ElementType | S11 | S22 | S12 | SMax | SMin | SVM | SAngle | S13 | S23 | SMaxOuter | SMinOuter | SVMOuter";
                case "Element Strains - Area Shells":
                    return "Format: Area | Joint | Output Case | Case Type | Step Type | Step Number | ElementType | E11 | E22 | E12 | EMax | EMin | EAngle | E13 | E23";
                case "Element Joint Forces - Shells":
                    return "Format: Area | Joint | Output Case | Case Type | Step Type | Step Number | ElementType | FX | FY | FZ | MX | MY | MZ";
                case "Pier Forces":
                    return "Format: Story | Pier | Output Case | Case Type | Step Type | Step Number | Location | P | V2 | V3 | T | M2 | M3";
                case "Objects and Elements - Joints":
                    return "Format: Joint | Label | Unique Name | Story | X-Coord | Y-Coord | Z-Coord";
                case "Objects and Elements - Frames":
                    return "Format: Frame | Label | Unique Name | Story | PointI | PointJ | Section | Material";
                case "Objects and Elements - Areas":
                    return "Format: Area | Label | Unique Name | Story | Section | Material";
                case "Point Object Connectivity":
                    return "Format: Point | Object Type | Object Name | Point Number";
                case "Beam Object Connectivity":
                case "Column Object Connectivity":
                case "Brace Object Connectivity":
                    return "Format: Object | Unique Name | PointI | PointJ";
                case "Floor Object Connectivity":
                case "Wall Object Connectivity":
                    return "Format: Object | Unique Name | Point1 | Point2 | Point3 | Point4";
                case "Tributary Area and LLRF":
                    return "Format: Story | Joint | Area | LLRF";
                case "Modal Periods And Frequencies":
                    return "Format: Case | Mode | Period | Frequency | CircFreq | Eigenvalue";
                case "Modal Participating Mass Ratios":
                    return "Format: Case | Mode | Period | UX | UY | UZ | SumUX | SumUY | SumUZ | RX | RY | RZ | SumRX | SumRY | SumRZ";
                case "Modal Load Participation Ratios":
                    return "Format: Case | ItemType | Item | Static | Dynamic";
                case "Modal Participation Factors":
                    return "Format: Case | Mode | Direction | Period | UX | UY | UZ | RX | RY | RZ | ModalMass | ModalStiff";
                case "Modal Direction Factors":
                    return "Format: Case | Mode | Period | UX | UY | UZ | RX | RY | RZ";
                case "Response Spectrum Modal Info":
                    return "Format: Output Case | Mode | Period | DX | DY | DZ | RX | RY | RZ | ModalMass | ModalStiff";
                case "Project Information":
                    return "Format: Key | Value";
                case "Material List by Object Type":
                    return "Format: Material | Object Type | NumPieces | TotalWeight";
                case "Material List by Section Property":
                    return "Format: Material | Section Property | NumPieces | TotalLength | TotalWeight";
                case "Material List by Story":
                    return "Format: Story | Material | NumPieces | TotalWeight";
                case "Mass Summary by Story":
                    return "Format: Story | Mass X | Mass Y | Mass Z | Center of Mass";
                case "Mass Summary by Diaphragm":
                    return "Format: Diaphragm | Story | Mass X | Mass Y | Mass Z";
                case "Mass Summary by Group":
                    return "Format: Group | Mass X | Mass Y | Mass Z";
                default:
                    return $"Export format for {tableName}";
            }
        }

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
