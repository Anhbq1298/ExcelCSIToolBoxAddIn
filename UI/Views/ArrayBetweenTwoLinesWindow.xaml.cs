using System;
using System.Windows;
using ExcelCSIToolBoxAddIn.UI.ViewModels;

namespace ExcelCSIToolBoxAddIn.UI.Views
{
    public partial class ArrayBetweenTwoLinesWindow : Window
    {
        public ArrayBetweenTwoLinesWindow(CsiToolboxViewModel viewModel)
        {
            InitializeComponent();
            DataContext = viewModel ?? throw new ArgumentNullException(nameof(viewModel));
        }
    }
}
