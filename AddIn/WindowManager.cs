using System;
using System.Windows.Controls;
using ExcelCSIToolBox.Application.UseCases;
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
        private static IExcelSelectionService _excelSelectionService;
        private static IExcelOutputService _excelOutputService;

        private static Microsoft.Office.Tools.CustomTaskPane _etabsPane;
        private static WpfTaskPaneHost _etabsHost;
        private static OutputTableExportOptionsWindow _getBaseReactionsWindow;
        private static GetModalMassParticipationRatiosWindow _modalMassParticipationRatiosWindow;
        private static GetStoryResultsWindow _storyForcesWindow;
        private static GetStoryResultsWindow _storyDriftsWindow;
        private static GetStoryResultsWindow _storyMaxOverAverageDisplacementsWindow;
        private static GetStoryResultsWindow _storyMaxOverAverageDriftsWindow;
        private static GetMassSummaryByStoryWindow _massSummaryByStoryWindow;

        internal static void Configure(
            ICSISapModelConnectionService etabsConnectionService,
            IExcelSelectionService excelSelectionService,
            IExcelOutputService excelOutputService)
        {
            _etabsConnectionService = etabsConnectionService ?? throw new ArgumentNullException(nameof(etabsConnectionService));
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


        internal static void ShowGetBaseReactionsWindow()
        {
            EnsureConfigured(_etabsConnectionService);

            if (_getBaseReactionsWindow != null)
            {
                _getBaseReactionsWindow.Activate();
                return;
            }

            var useCases = new CsiToolboxUseCaseBundle(_etabsConnectionService, _excelSelectionService, _excelOutputService);
            var config = new OutputTableExportConfig
            {
                TableDisplayName = "Base Reactions",
                Breadcrumb = "ETABS Toolbox / ANALYSIS RESULTS / Base Reactions",
                Description = "Select output cases to export Base Reactions.",
                PopupProfileKey = "ForceOutput"
            };
            var window = OutputTableExportWorkflow.Run(
                config,
                useCases,
                _etabsConnectionService,
                _excelOutputService);
            window.Closed += delegate { _getBaseReactionsWindow = null; };
            _getBaseReactionsWindow = window;
        }

        internal static void ShowModalMassParticipationRatiosWindow()
        {
            EnsureConfigured(_etabsConnectionService);

            if (_modalMassParticipationRatiosWindow != null)
            {
                _modalMassParticipationRatiosWindow.Activate();
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
            window.Show();
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
                _massSummaryByStoryWindow.Activate();
                return;
            }

            var useCases = new CsiToolboxUseCaseBundle(_etabsConnectionService, _excelSelectionService, _excelOutputService);
            var viewModel = new GetMassSummaryByStoryViewModel(useCases, _etabsConnectionService, _excelOutputService);
            var window = new GetMassSummaryByStoryWindow(viewModel);
            window.Closed += delegate { _massSummaryByStoryWindow = null; };
            _massSummaryByStoryWindow = window;
            window.Show();
        }

        private static void ShowStoryResultsWindow(
            StoryPostprocessingResultKind kind,
            ref GetStoryResultsWindow existingWindow)
        {
            EnsureConfigured(_etabsConnectionService);

            if (existingWindow != null)
            {
                existingWindow.Activate();
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

            window.Show();
        }

        internal static void DisposePanes()
        {
            DisposePane(ref _etabsPane, ref _etabsHost);
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
                var useCases = new CsiToolboxUseCaseBundle(connectionService, _excelSelectionService, _excelOutputService);
                control.DataContext = new CsiToolboxViewModel(
                    useCases,
                    connectionService,
                    _excelSelectionService,
                    _excelOutputService);

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
