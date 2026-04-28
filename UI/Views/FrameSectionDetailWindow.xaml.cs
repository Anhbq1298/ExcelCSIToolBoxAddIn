using System.Windows;
using ExcelCSIToolBoxAddIn.Data.DTOs;
using ExcelCSIToolBoxAddIn.UI.Helpers;

namespace ExcelCSIToolBoxAddIn.UI.Views
{
    public partial class FrameSectionDetailWindow : Window
    {
        public FrameSectionDetailWindow(object dataContext)
        {
            InitializeComponent();
            DataContext = dataContext;
            Loaded += FrameSectionDetailWindow_Loaded;
        }

        private void FrameSectionDetailWindow_Loaded(object sender, RoutedEventArgs e)
        {
            if (DataContext is CSISapModelFrameSectionDetailDTO dto)
                SectionShapeRenderer.Render(PreviewCanvas, dto);
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
