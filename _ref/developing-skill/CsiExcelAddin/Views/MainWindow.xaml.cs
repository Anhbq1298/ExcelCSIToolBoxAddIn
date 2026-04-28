using System.Windows;
using CsiExcelAddin.ViewModels;

namespace CsiExcelAddin.Views
{
    /// <summary>
    /// Code-behind for MainWindow.
    /// Responsibility is limited to receiving the ViewModel from the composition root
    /// and assigning it as DataContext. No business logic lives here.
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow(MainViewModel viewModel)
        {
            InitializeComponent();

            // Assign ViewModel as DataContext so all XAML bindings resolve
            DataContext = viewModel;
        }
    }
}

