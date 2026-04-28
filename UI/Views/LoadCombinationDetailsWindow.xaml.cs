using System.Collections.Generic;
using System.Windows;
using ExcelCSIToolBoxAddIn.Data.DTOs;

namespace ExcelCSIToolBoxAddIn.UI.Views
{
    public partial class LoadCombinationDetailsWindow : Window
    {
        public LoadCombinationDetailsWindow(IReadOnlyList<LoadCombinationItemDTO> items)
        {
            InitializeComponent();
            DataContext = items;
        }
    }
}
