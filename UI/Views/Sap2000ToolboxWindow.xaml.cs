using System.Windows;

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
