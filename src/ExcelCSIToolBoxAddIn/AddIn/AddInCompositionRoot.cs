using System;
using ExcelCSIToolBox.Application.GenerativeDesign;
using ExcelCSIToolBox.Application.ToolCatalog;
using ExcelCSIToolBox.Application.ToolCatalog.Contracts;
using ExcelCSIToolBox.Core.Abstractions;
using ExcelCSIToolBox.AI.Agent;
using ExcelCSIToolBox.AI.Mcp.Client;
using ExcelCSIToolBox.AI.Mcp.Safety;
using ExcelCSIToolBox.AI.Mcp.Server;
using ExcelCSIToolBox.AI.Providers.Ollama;
using ExcelCSIToolBox.Core.Abstractions.CSI;
using ExcelCSIToolBox.Core.Abstractions.Excel;
using ExcelCSIToolBox.Infrastructure.CSI.Common;
using ExcelCSIToolBox.Infrastructure.CSI.Common.Modelling.Random;
using ExcelCSIToolBox.Infrastructure.CSI.Common.ReadOnly;
using ExcelCSIToolBox.Infrastructure.CSI.Common.Modelling.Truss;
using ExcelCSIToolBox.Infrastructure.CSI.Common.Workflow;
using ExcelCSIToolBox.Infrastructure.Excel.Interop;
using ExcelCSIToolBoxAddIn.UI.ViewModels;
using ExcelCSIToolBoxAddIn.UI.Views;

namespace ExcelCSIToolBoxAddIn.AddIn
{
    internal static class AddInCompositionRoot
    {
        private static ICSISapModelConnectionService _etabsConnectionService;
        private static ICSISapModelConnectionService _sap2000ConnectionService;
        private static IExcelSelectionService _excelSelectionService;
        private static IExcelOutputService _excelOutputService;
        private static IToolCatalogService _toolCatalogService;
        private static IProgressReporter _progressReporter;
        private static IMutationGuard _mutationGuard;
        private static IThreadDispatcher _threadDispatcher;

        public static void Configure(
            ICSISapModelConnectionService etabsConnectionService,
            ICSISapModelConnectionService sap2000ConnectionService,
            IProgressReporter progressReporter,
            IThreadDispatcher threadDispatcher)
        {
            _etabsConnectionService = etabsConnectionService ?? throw new ArgumentNullException(nameof(etabsConnectionService));
            _sap2000ConnectionService = sap2000ConnectionService ?? throw new ArgumentNullException(nameof(sap2000ConnectionService));
            _progressReporter = progressReporter ?? throw new ArgumentNullException(nameof(progressReporter));
            _threadDispatcher = threadDispatcher ?? throw new ArgumentNullException(nameof(threadDispatcher));
            _mutationGuard = new WpfMutationGuard();
            _excelSelectionService = new ExcelSelectionService();
            _excelOutputService = new ExcelOutputService();
            _toolCatalogService = new ToolCatalogService(_etabsConnectionService, _sap2000ConnectionService);

            WindowManager.Configure(
                _etabsConnectionService,
                _sap2000ConnectionService,
                _excelSelectionService,
                _excelOutputService);

            AiTaskPaneManager.Configure(CreateAiAgentChatControl);
        }

        public static AiAgentChatControl CreateAiAgentChatControl()
        {
            IAiChatSessionService sessionService = CreateAiChatSessionService();
            return new AiAgentChatControl(
                new AiAgentChatViewModel(sessionService, _threadDispatcher),
                _threadDispatcher);
        }

        private static IAiChatSessionService CreateAiChatSessionService()
        {
            EnsureConfigured();

            var writeGuard = new CsiWriteGuard();
            var operationLogger = new CsiOperationLogger();
            var commandService = new CsiModelCommandService(
                _etabsConnectionService,
                _sap2000ConnectionService,
                writeGuard,
                operationLogger);

            var context = new CsiMcpToolContext(
                new CsiReadOnlyConnectionService(),
                new CsiReadOnlySelectionService(),
                new CsiReadOnlyFrameService(),
                _etabsConnectionService,
                _sap2000ConnectionService,
                commandService,
                writeGuard,
                operationLogger,
                new CsiRandomObjectGenerationService(),
                new CsiHoweTrussGenerationService(),
                new CsiWorkflowExecutionService(),
                _toolCatalogService,
                _mutationGuard,
                new BuildingOptionService(),
                new ConstraintValidationService(),
                new ResultEvaluationService(),
                new OptionRankingService());

            var mcpServer = new LocalMcpServer(context);
            var mcpClient = new LocalMcpClient(mcpServer);
            var ollamaService = new OllamaChatService();
            var orchestrator = new AiAgentOrchestrator(ollamaService, mcpClient);

            return new AiChatSessionService(
                orchestrator,
                _etabsConnectionService,
                _sap2000ConnectionService,
                AiModelDefaults.DefaultOllamaModel);
        }

        private static void EnsureConfigured()
        {
            if (_etabsConnectionService == null || _sap2000ConnectionService == null)
            {
                throw new InvalidOperationException("The add-in composition root is not configured.");
            }
        }

        public static void RefreshPlugin()
        {
            try
            {
                string tempVbsPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "reload_excel_csi_addin.vbs");
                string vbsScript = @"
WScript.Sleep 1000
On Error Resume Next
Dim excel, addin, i
Set excel = Nothing
For i = 1 To 10
    Err.Clear
    Set excel = GetObject(, ""Excel.Application"")
    If Err.Number = 0 And Not excel Is Nothing Then
        Exit For
    End If
    WScript.Sleep 500
Next
If Not excel Is Nothing Then
    Set addin = Nothing
    For i = 1 To 10
        Err.Clear
        Set addin = excel.COMAddIns(""ExcelCSIToolBoxAddIn"")
        If Err.Number = 0 And Not addin Is Nothing Then
            Exit For
        End If
        WScript.Sleep 500
    Next
    If Not addin Is Nothing Then
        For i = 1 To 10
            Err.Clear
            addin.Connect = False
            If Err.Number = 0 Then
                Exit For
            End If
            WScript.Sleep 500
        Next
        WScript.Sleep 1000
        For i = 1 To 10
            Err.Clear
            addin.Connect = True
            If Err.Number = 0 Then
                Exit For
            End If
            WScript.Sleep 500
        Next
    End If
End If
";
                System.IO.File.WriteAllText(tempVbsPath, vbsScript, System.Text.Encoding.ASCII);
                System.Diagnostics.Process.Start("wscript.exe", "\"" + tempVbsPath + "\"");
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(
                    "Failed to restart the add-in: " + ex.Message,
                    "Refresh Plugin Error",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error);
            }
        }
    }
}
