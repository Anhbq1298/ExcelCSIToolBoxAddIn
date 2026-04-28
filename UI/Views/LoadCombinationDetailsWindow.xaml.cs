using System.Collections.Generic;
using System.Windows;
using ExcelCSIToolBox.Data.DTOs.CSI;

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

