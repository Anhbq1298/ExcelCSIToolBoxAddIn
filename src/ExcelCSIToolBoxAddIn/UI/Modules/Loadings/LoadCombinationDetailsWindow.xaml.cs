using System.Collections.Generic;
using System.Windows;
using ExcelCSIToolBox.Core.Contracts.CSI;

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

