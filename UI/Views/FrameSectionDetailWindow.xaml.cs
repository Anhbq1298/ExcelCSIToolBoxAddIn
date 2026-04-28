using System.Globalization;
using System.Windows;
using ExcelCSIToolBoxAddIn.Data.DTOs;
using ExcelCSIToolBoxAddIn.UI.Helpers;
using ExcelCSIToolBoxAddIn.UI.ViewModels;

namespace ExcelCSIToolBoxAddIn.UI.Views
{
    public partial class FrameSectionDetailWindow : Window
    {
        private readonly CSISapModelFrameSectionDetailDTO _previewDetail;

        public FrameSectionDetailWindow(CSISapModelFrameSectionDetailDTO detail)
        {
            InitializeComponent();
            _previewDetail = detail;
            ViewModel = new FrameSectionDetailViewModel(detail);
            DataContext = ViewModel;
            Loaded += FrameSectionDetailWindow_Loaded;
        }

        public FrameSectionDetailViewModel ViewModel { get; }

        private void FrameSectionDetailWindow_Loaded(object sender, RoutedEventArgs e)
        {
            SectionShapeRenderer.Render(PreviewCanvas, _previewDetail);
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(ViewModel.SectionName))
            {
                MessageBox.Show("Section name is required.", "Section Property", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (string.IsNullOrWhiteSpace(ViewModel.MaterialName))
            {
                MessageBox.Show("Material name is required.", "Section Property", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            foreach (var dimension in ViewModel.Dimensions)
            {
                if (!double.TryParse(dimension.ValueText, NumberStyles.Float, CultureInfo.InvariantCulture, out double value) &&
                    !double.TryParse(dimension.ValueText, NumberStyles.Float, CultureInfo.CurrentCulture, out value))
                {
                    MessageBox.Show($"Invalid value for {dimension.Key}.", "Section Property", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                if (value <= 0)
                {
                    MessageBox.Show($"{dimension.Key} must be greater than zero.", "Section Property", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
            }

            DialogResult = true;
            Close();
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}
