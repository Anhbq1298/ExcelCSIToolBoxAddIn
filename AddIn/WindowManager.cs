using System;
using System.Windows;
using ExcelCSIToolBox.Core.Abstractions.CSI;
using ExcelCSIToolBox.Core.Abstractions.Excel;
using ExcelCSIToolBox.Infrastructure.CSISapModel;
using ExcelCSIToolBox.Infrastructure.Excel;
using ExcelCSIToolBoxAddIn.UI.ViewModels;
using ExcelCSIToolBoxAddIn.UI.Views;

namespace ExcelCSIToolBoxAddIn.AddIn
{
    internal static class WindowManager
    {
        private static ICSISapModelConnectionService _etabsConnectionService;
        private static ICSISapModelConnectionService _sap2000ConnectionService;
        private static IExcelSelectionService _excelSelectionService;
        private static IExcelOutputService _excelOutputService;
        private static EtabsToolboxWindow _activeEtabsWindow;
        private static Sap2000ToolboxWindow _activeSap2000Window;

        internal static void Configure(
            ICSISapModelConnectionService etabsConnectionService,
            ICSISapModelConnectionService sap2000ConnectionService,
            IExcelSelectionService excelSelectionService,
            IExcelOutputService excelOutputService)
        {
            _etabsConnectionService = etabsConnectionService ?? throw new ArgumentNullException(nameof(etabsConnectionService));
            _sap2000ConnectionService = sap2000ConnectionService ?? throw new ArgumentNullException(nameof(sap2000ConnectionService));
            _excelSelectionService = excelSelectionService ?? throw new ArgumentNullException(nameof(excelSelectionService));
            _excelOutputService = excelOutputService ?? throw new ArgumentNullException(nameof(excelOutputService));
        }

        internal static void ShowEtabsWindow()
        {
            ShowCsiWindow(
                _etabsConnectionService,
                () => _activeEtabsWindow,
                window => _activeEtabsWindow = window);
        }

        internal static void ShowSap2000Window()
        {
            ShowCsiWindow(
                _sap2000ConnectionService,
                () => _activeSap2000Window,
                window => _activeSap2000Window = window);
        }

        private static void ShowCsiWindow<TWindow>(
            ICSISapModelConnectionService connectionService,
            Func<TWindow> getActiveWindow,
            Action<TWindow> setActiveWindow)
            where TWindow : Window, new()
        {
            if (connectionService == null || _excelSelectionService == null || _excelOutputService == null)
            {
                throw new InvalidOperationException("WindowManager is not configured.");
            }

            var activeWindow = getActiveWindow();
            if (activeWindow == null)
            {
                var viewModel = new CsiToolboxViewModel(connectionService, _excelSelectionService, _excelOutputService);
                activeWindow = new TWindow
                {
                    DataContext = viewModel
                };
                setActiveWindow(activeWindow);

                var openedWindow = activeWindow;
                activeWindow.Closed += (_, __) =>
                {
                    if (ReferenceEquals(getActiveWindow(), openedWindow))
                    {
                        setActiveWindow(null);
                    }
                };

                System.Windows.Forms.Integration.ElementHost.EnableModelessKeyboardInterop(activeWindow);
                activeWindow.Show();
                activeWindow.Activate();
                activeWindow.Focus();
                return;
            }

            if (activeWindow.WindowState == WindowState.Minimized)
            {
                activeWindow.WindowState = WindowState.Normal;
            }

            if (!activeWindow.IsVisible)
            {
                activeWindow.Show();
            }

            activeWindow.Topmost = true;
            activeWindow.Topmost = false;
            activeWindow.Activate();
            activeWindow.Focus();
        }
    }
}

