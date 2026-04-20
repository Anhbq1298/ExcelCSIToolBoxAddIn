using System;
using System.Windows;
using ExcelCSIToolBoxAddIn.Infrastructure.Etabs;
using ExcelCSIToolBoxAddIn.Infrastructure.Excel;
using ExcelCSIToolBoxAddIn.UI.ViewModels;
using ExcelCSIToolBoxAddIn.UI.Views;

namespace ExcelCSIToolBoxAddIn.AddIn
{
    /// <summary>
    /// Opens the ETABS toolbox shell window from the add-in layer.
    /// Keeps window creation logic out of ribbon click handlers.
    /// </summary>
    public class EtabsToolboxWindowLauncher
    {
        private readonly IEtabsConnectionService _etabsConnectionService;
        private readonly IExcelSelectionService _excelSelectionService;
        private readonly IExcelOutputService _excelOutputService;

        public EtabsToolboxWindowLauncher(
            IEtabsConnectionService etabsConnectionService,
            IExcelSelectionService excelSelectionService,
            IExcelOutputService excelOutputService)
        {
            _etabsConnectionService = etabsConnectionService ?? throw new ArgumentNullException(nameof(etabsConnectionService));
            _excelSelectionService = excelSelectionService ?? throw new ArgumentNullException(nameof(excelSelectionService));
            _excelOutputService = excelOutputService ?? throw new ArgumentNullException(nameof(excelOutputService));
        }

        public void OpenWindow()
        {
            var viewModel = new EtabsToolboxViewModel(_etabsConnectionService, _excelSelectionService, _excelOutputService);
            var window = new EtabsToolboxWindow
            {
                DataContext = viewModel
            };

            // Show as modeless so users can continue working in Excel.
            window.Show();
            window.Activate();
        }
    }
}
