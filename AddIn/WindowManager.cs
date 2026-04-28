using System;
using System.Windows.Controls;
using ExcelCSIToolBox.Core.Abstractions.CSI;
using ExcelCSIToolBox.Core.Abstractions.Excel;
using ExcelCSIToolBoxAddIn.UI.ViewModels;
using ExcelCSIToolBoxAddIn.UI.Views;
using Microsoft.Office.Core;

namespace ExcelCSIToolBoxAddIn.AddIn
{
    internal static class WindowManager
    {
        private static ICSISapModelConnectionService _etabsConnectionService;
        private static ICSISapModelConnectionService _sap2000ConnectionService;
        private static IExcelSelectionService _excelSelectionService;
        private static IExcelOutputService _excelOutputService;

        private static Microsoft.Office.Tools.CustomTaskPane _etabsPane;
        private static Microsoft.Office.Tools.CustomTaskPane _sap2000Pane;
        private static WpfTaskPaneHost _etabsHost;
        private static WpfTaskPaneHost _sap2000Host;

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
            ToggleCsiPane(
                ref _etabsPane,
                ref _etabsHost,
                "ETABS Toolbox",
                _etabsConnectionService,
                () => new EtabsToolboxControl());
        }

        internal static void ShowSap2000Window()
        {
            ToggleCsiPane(
                ref _sap2000Pane,
                ref _sap2000Host,
                "SAP2000 Toolbox",
                _sap2000ConnectionService,
                () => new Sap2000ToolboxControl());
        }

        internal static void DisposePanes()
        {
            DisposePane(ref _etabsPane, ref _etabsHost);
            DisposePane(ref _sap2000Pane, ref _sap2000Host);
        }

        private static void ToggleCsiPane(
            ref Microsoft.Office.Tools.CustomTaskPane pane,
            ref WpfTaskPaneHost host,
            string title,
            ICSISapModelConnectionService connectionService,
            Func<UserControl> createControl)
        {
            EnsureConfigured(connectionService);

            if (pane == null)
            {
                UserControl control = createControl();
                control.DataContext = new CsiToolboxViewModel(connectionService, _excelSelectionService, _excelOutputService);

                host = new WpfTaskPaneHost(control);
                pane = Globals.ExcelCSIToolBoxAddin.CustomTaskPanes.Add(host, title);
                pane.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight;
                pane.Width = 820;
            }

            pane.Visible = !pane.Visible;
        }

        private static void EnsureConfigured(ICSISapModelConnectionService connectionService)
        {
            if (Globals.ExcelCSIToolBoxAddin == null)
            {
                throw new InvalidOperationException("The Excel add-in is not initialized.");
            }

            if (connectionService == null || _excelSelectionService == null || _excelOutputService == null)
            {
                throw new InvalidOperationException("WindowManager is not configured.");
            }
        }

        private static void DisposePane(
            ref Microsoft.Office.Tools.CustomTaskPane pane,
            ref WpfTaskPaneHost host)
        {
            if (pane != null)
            {
                Globals.ExcelCSIToolBoxAddin.CustomTaskPanes.Remove(pane);
                pane = null;
            }

            if (host != null)
            {
                host.Dispose();
                host = null;
            }
        }
    }
}
