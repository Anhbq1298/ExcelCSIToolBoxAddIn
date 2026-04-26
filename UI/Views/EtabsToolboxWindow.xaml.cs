using System.Windows;

namespace ExcelCSIToolBoxAddIn.UI.Views
{
    /// <summary>
    /// Interaction logic for EtabsToolboxWindow.xaml
    /// </summary>
    public partial class EtabsToolboxWindow : Window
    {
        public EtabsToolboxWindow()
        {
            InitializeComponent();
        }

        private void HelpButton_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show(
                "This tool was developed by Mark Bui Quang Anh.",
                "About ETABS Toolbox",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
        }
    }
}
