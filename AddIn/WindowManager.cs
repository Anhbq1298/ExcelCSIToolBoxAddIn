using System;
using System.Windows;
using ExcelCSIToolBoxAddIn.Infrastructure.Etabs;
using ExcelCSIToolBoxAddIn.Infrastructure.Excel;
using ExcelCSIToolBoxAddIn.UI.ViewModels;
using ExcelCSIToolBoxAddIn.UI.Views;

namespace ExcelCSIToolBoxAddIn.AddIn
{
    internal static class WindowManager
    {
        private static IEtabsConnectionService _etabsConnectionService;
        private static IExcelSelectionService _excelSelectionService;
        private static IExcelOutputService _excelOutputService;
        private static EtabsToolboxWindow _activeEtabsWindow;

        internal static void Configure(
            IEtabsConnectionService etabsConnectionService,
            IExcelSelectionService excelSelectionService,
            IExcelOutputService excelOutputService)
        {
            _etabsConnectionService = etabsConnectionService ?? throw new ArgumentNullException(nameof(etabsConnectionService));
            _excelSelectionService = excelSelectionService ?? throw new ArgumentNullException(nameof(excelSelectionService));
            _excelOutputService = excelOutputService ?? throw new ArgumentNullException(nameof(excelOutputService));
        }

        internal static void ShowEtabsWindow()
        {
            if (_etabsConnectionService == null || _excelSelectionService == null || _excelOutputService == null)
            {
                throw new InvalidOperationException("WindowManager is not configured.");
            }

            if (_activeEtabsWindow == null)
            {
                var viewModel = new EtabsToolboxViewModel(_etabsConnectionService, _excelSelectionService, _excelOutputService);
                _activeEtabsWindow = new EtabsToolboxWindow
                {
                    DataContext = viewModel
                };

                _activeEtabsWindow.Closed += (_, __) => _activeEtabsWindow = null;
                _activeEtabsWindow.Show();
                _activeEtabsWindow.Activate();
                _activeEtabsWindow.Focus();
                return;
            }

            if (_activeEtabsWindow.WindowState == WindowState.Minimized)
            {
                _activeEtabsWindow.WindowState = WindowState.Normal;
            }

            if (!_activeEtabsWindow.IsVisible)
            {
                _activeEtabsWindow.Show();
            }

            _activeEtabsWindow.Topmost = true;
            _activeEtabsWindow.Topmost = false;
            _activeEtabsWindow.Activate();
            _activeEtabsWindow.Focus();
        }
    }
}
