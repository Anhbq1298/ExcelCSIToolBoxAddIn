using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;
using ExcelCSIToolBox.Application.Composition;
using ExcelCSIToolBox.Core.Abstractions.CSI;
using ExcelCSIToolBox.Core.Abstractions.Excel;
using ExcelCSIToolBox.Core.Common.Results;
using ExcelCSIToolBoxAddIn.AddIn.Composition;
using ExcelCSIToolBoxAddIn.AddIn.Diagnostics;
using ExcelCSIToolBoxAddIn.UI.ViewModels;
using ExcelCSIToolBoxAddIn.UI.Views;
using Microsoft.Office.Core;

namespace ExcelCSIToolBoxAddIn.AddIn
{
    internal static class WindowManager
    {
        private const int CsiPaneWidth = 560;

        private static ICSISapModelConnectionService _etabsConnectionService;
        private static ICSISapModelConnectionService _sap2000ConnectionService;
        private static IExcelSelectionService _excelSelectionService;
        private static IExcelOutputService _excelOutputService;

        private static Microsoft.Office.Tools.CustomTaskPane _etabsPane;
        private static WpfTaskPaneHost _etabsHost;
        private static Microsoft.Office.Tools.CustomTaskPane _sap2000Pane;
        private static WpfTaskPaneHost _sap2000Host;
        private static GetModalMassParticipationRatiosWindow _modalMassParticipationRatiosWindow;
        private static GetStoryResultsWindow _storyForcesWindow;
        private static GetStoryResultsWindow _storyDriftsWindow;
        private static GetStoryResultsWindow _storyMaxOverAverageDisplacementsWindow;
        private static GetStoryResultsWindow _storyMaxOverAverageDriftsWindow;
        private static GetMassSummaryByStoryWindow _massSummaryByStoryWindow;
        private static AboutWindow _aboutWindow;
        private static DropPanelWindow _dropPanelWindow;


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
            ShowCsiPane(
                ref _etabsPane,
                ref _etabsHost,
                "ETABS Toolbox",
                _etabsConnectionService,
                () => new EtabsToolboxControl());
        }

        internal static void ShowSap2000Window()
        {
            ShowCsiPane(
                ref _sap2000Pane,
                ref _sap2000Host,
                "SAP2000 Toolbox",
                _sap2000ConnectionService,
                () => new EtabsToolboxControl());
        }


        internal static void ShowGetBaseReactionsWindow()
        {
            EnsureConfigured(_etabsConnectionService);

            var useCases = new CsiToolboxUseCaseBundle(_etabsConnectionService, _excelSelectionService, _excelOutputService);
            var config = new OutputTableExportConfig
            {
                TableDisplayName = "Base Reactions",
                Breadcrumb = "ETABS Toolbox / ANALYSIS RESULTS / Base Reactions",
                Description = "Select output cases to export Base Reactions.",
                PopupProfileKey = "ForceOutput"
            };
            OutputTableExportWorkflow.Run(
                config,
                useCases,
                _etabsConnectionService,
                _excelOutputService);
        }

        internal static void ShowModalMassParticipationRatiosWindow()
        {
            EnsureConfigured(_etabsConnectionService);

            if (_modalMassParticipationRatiosWindow != null)
            {
                ModelessWpfWindowService.Show(_modalMassParticipationRatiosWindow);
                return;
            }

            var useCases = new CsiToolboxUseCaseBundle(_etabsConnectionService, _excelSelectionService, _excelOutputService);
            var viewModel = new GetModalMassParticipationRatiosViewModel(
                useCases,
                _etabsConnectionService,
                _excelOutputService);
            var window = new GetModalMassParticipationRatiosWindow(viewModel);
            window.Closed += delegate { _modalMassParticipationRatiosWindow = null; };
            _modalMassParticipationRatiosWindow = window;
            ModelessWpfWindowService.Show(window);
        }

        internal static void ShowStoryForcesWindow()
        {
            EnsureConfigured(_etabsConnectionService);
            var useCases = new CsiToolboxUseCaseBundle(_etabsConnectionService, _excelSelectionService, _excelOutputService);
            var config = new OutputTableExportConfig
            {
                TableDisplayName = "Story Forces",
                Breadcrumb = "ETABS Toolbox / ANALYSIS RESULTS / Other Output Items / Story Forces",
                Description = "Select load case or load combination and output unit to export Story Forces.",
                PopupProfileKey = "StoryForces"
            };
            OutputTableExportWorkflow.Run(config, useCases, _etabsConnectionService, _excelOutputService);
        }

        internal static void ShowStoryDriftsWindow()
        {
            ShowStoryResultsWindow(
                StoryPostprocessingResultKind.StoryDrifts,
                ref _storyDriftsWindow);
        }

        internal static void ShowStoryMaxOverAverageDisplacementsWindow()
        {
            ShowStoryResultsWindow(
                StoryPostprocessingResultKind.StoryMaxOverAverageDisplacements,
                ref _storyMaxOverAverageDisplacementsWindow);
        }

        internal static void ShowStoryMaxOverAverageDriftsWindow()
        {
            ShowStoryResultsWindow(
                StoryPostprocessingResultKind.StoryMaxOverAverageDrifts,
                ref _storyMaxOverAverageDriftsWindow);
        }

        internal static void ShowMassSummaryByStoryWindow()
        {
            EnsureConfigured(_etabsConnectionService);
            if (_massSummaryByStoryWindow != null)
            {
                ModelessWpfWindowService.Show(_massSummaryByStoryWindow);
                return;
            }

            var useCases = new CsiToolboxUseCaseBundle(_etabsConnectionService, _excelSelectionService, _excelOutputService);
            var viewModel = new GetMassSummaryByStoryViewModel(useCases, _etabsConnectionService, _excelOutputService);
            var window = new GetMassSummaryByStoryWindow(viewModel);
            window.Closed += delegate { _massSummaryByStoryWindow = null; };
            _massSummaryByStoryWindow = window;
            ModelessWpfWindowService.Show(window);
        }

        internal static void ShowAboutWindow()
        {
            if (_aboutWindow != null)
            {
                ModelessWpfWindowService.Show(_aboutWindow);
                return;
            }

            var window = new AboutWindow();
            window.Closed += delegate { _aboutWindow = null; };
            _aboutWindow = window;
            ModelessWpfWindowService.Show(window);
        }

        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr windowHandle);

        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr windowHandle, int command);

        private const int ShowWindowRestore = 9;

        private static void ActivateConnectedCsiWindow()
        {
            try
            {
                OperationResult<ExcelCSIToolBox.Core.Contracts.CSI.CSISapModelConnectionInfoDTO> connectionResult =
                    _etabsConnectionService.GetCurrentConnection();
                if (!connectionResult.IsSuccess ||
                    connectionResult.Data == null ||
                    !connectionResult.Data.ProcessId.HasValue)
                {
                    return;
                }

                Process process = Process.GetProcessById(connectionResult.Data.ProcessId.Value);
                if (process.MainWindowHandle == IntPtr.Zero)
                {
                    return;
                }

                ShowWindow(process.MainWindowHandle, ShowWindowRestore);
                SetForegroundWindow(process.MainWindowHandle);
            }
            catch
            {
                // Selection polling remains available if native window activation is unavailable.
            }
        }

        private static Window GetActiveOwnerWindow()
        {
            if (System.Windows.Application.Current == null)
            {
                return null;
            }

            foreach (Window window in System.Windows.Application.Current.Windows)
            {
                if (window != null && window.IsActive)
                {
                    return window;
                }
            }

            return System.Windows.Application.Current.MainWindow;
        }

        internal static void ShowDropPanelWindow()
        {
            EnsureConfigured(_etabsConnectionService);
            if (_dropPanelWindow != null)
            {
                ModelessWpfWindowService.Show(_dropPanelWindow);
                return;
            }

            var dropPanelService = AppServiceFactory.CreateDropPanelService(_etabsConnectionService);
            var settingsStore = new DropPanelSettingsStore();
            var logExporter = new DropPanelExcelLogExporter();
            
            DropPanelWindow window = null;
            DropPanelViewModel viewModel = null;
            viewModel = new DropPanelViewModel(
                _etabsConnectionService,
                dropPanelService,
                settingsStore,
                logExporter,
                delegate { return window; },
                delegate
                {
                    if (window != null)
                    {
                        window.Close();
                    }
                });

            window = new DropPanelWindow(viewModel);
            window.Closed += delegate { _dropPanelWindow = null; };
            _dropPanelWindow = window;
            ModelessWpfWindowService.Show(window);
        }

        private static void ShowStoryResultsWindow(
            StoryPostprocessingResultKind kind,
            ref GetStoryResultsWindow existingWindow)
        {
            EnsureConfigured(_etabsConnectionService);

            if (existingWindow != null)
            {
                ModelessWpfWindowService.Show(existingWindow);
                return;
            }

            var useCases = new CsiToolboxUseCaseBundle(_etabsConnectionService, _excelSelectionService, _excelOutputService);
            var viewModel = new GetStoryResultsViewModel(
                kind,
                useCases,
                _etabsConnectionService,
                _excelOutputService);
            var window = new GetStoryResultsWindow(viewModel);

            if (kind == StoryPostprocessingResultKind.StoryForces)
            {
                window.Closed += delegate { _storyForcesWindow = null; };
                _storyForcesWindow = window;
            }
            else if (kind == StoryPostprocessingResultKind.StoryDrifts)
            {
                window.Closed += delegate { _storyDriftsWindow = null; };
                _storyDriftsWindow = window;
            }
            else if (kind == StoryPostprocessingResultKind.StoryMaxOverAverageDisplacements)
            {
                window.Closed += delegate { _storyMaxOverAverageDisplacementsWindow = null; };
                _storyMaxOverAverageDisplacementsWindow = window;
            }
            else
            {
                window.Closed += delegate { _storyMaxOverAverageDriftsWindow = null; };
                _storyMaxOverAverageDriftsWindow = window;
            }

            ModelessWpfWindowService.Show(window);
        }

        internal static void DisposePanes()
        {
            DisposePane(ref _etabsPane, ref _etabsHost);
            DisposePane(ref _sap2000Pane, ref _sap2000Host);
        }

        private static void ShowCsiPane(
            ref Microsoft.Office.Tools.CustomTaskPane pane,
            ref WpfTaskPaneHost host,
            string title,
            ICSISapModelConnectionService connectionService,
            Func<UserControl> createControl)
        {
            EnsureConfigured(connectionService);
            CollapseExpandedFormulaBar();

            if (pane == null)
            {
                AddInDiagnostics.Log("Creating task pane: " + title + ".");
                UserControl control = createControl();
                AddInDiagnostics.Log("Created WPF control for task pane: " + title + ".");
                var useCases = new CsiToolboxUseCaseBundle(connectionService, _excelSelectionService, _excelOutputService);
                var analysisResultServices = AppServiceFactory.CreateAnalysisResultServices(connectionService);
                control.DataContext = new CsiToolboxViewModel(
                    useCases,
                    connectionService,
                    _excelSelectionService,
                    _excelOutputService,
                    analysisResultServices);
                AddInDiagnostics.Log("Assigned DataContext for task pane: " + title + ".");

                host = new WpfTaskPaneHost(control);
                pane = Globals.ExcelCSIToolBoxAddin.CustomTaskPanes.Add(host, title);
                ApplyCsiPaneLayout(pane, title);
                AddInDiagnostics.Log("Task pane created: " + title + ".");
            }

            ApplyCsiPaneLayout(pane, title);
            pane.Visible = true;
            ApplyCsiPaneLayout(pane, title);
            AddInDiagnostics.Log("Task pane visible: " + title + ".");
        }

        private static void CollapseExpandedFormulaBar()
        {
            try
            {
                var application = Globals.ExcelCSIToolBoxAddin.Application;
                if (application != null && application.FormulaBarHeight > 1)
                {
                    application.FormulaBarHeight = 1;
                    AddInDiagnostics.Log("Excel formula bar height collapsed to 1.");
                }
            }
            catch (Exception ex)
            {
                AddInDiagnostics.Log("Could not collapse Excel formula bar: " + ex.Message);
            }
        }

        private static void ApplyCsiPaneLayout(Microsoft.Office.Tools.CustomTaskPane pane, string title)
        {
            pane.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight;
            pane.DockPositionRestrict = MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoHorizontal;
            pane.Width = CsiPaneWidth;
            AddInDiagnostics.Log(
                "Task pane layout applied: " + title +
                ", DockPosition=" + pane.DockPosition +
                ", Width=" + pane.Width + ".");
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
