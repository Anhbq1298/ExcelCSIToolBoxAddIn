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
using ExcelCSIToolBox.Infrastructure.Excel.Reading;
using ExcelCSIToolBox.Infrastructure.Excel.Writing;
using ExcelCSIToolBoxAddIn.AddIn.Diagnostics;
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
            AddInDiagnostics.Log("Refresh Plugin requested. Hot reload was not attempted because unloading a running VSTO add-in can crash Excel.");
            System.Windows.Forms.MessageBox.Show(
                "Excel cannot safely reload this VSTO add-in while it is running.\r\n\r\n" +
                "Save your work, close all Excel windows, and reopen Excel to load the latest build.",
                "Restart Excel to Refresh Plugin",
                System.Windows.Forms.MessageBoxButtons.OK,
                System.Windows.Forms.MessageBoxIcon.Information);
        }

        // Retained temporarily for reference while the refresh workflow is replaced.
        // Do not call this from Excel: changing COMAddIn.Connect for the currently
        // executing VSTO add-in has caused native access violations in mso.dll.
        private static void RefreshPluginUsingComReconnect()
        {
            try
            {
                AddInDiagnostics.Log("Refresh Plugin requested.");
                CloseModelessWpfWindowsForRefresh();
                WindowManager.DisposePanes();
                AiTaskPaneManager.DisposePane();

                string tempVbsPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "reload_excel_csi_addin.vbs");
                string logFolder = System.IO.Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "ExcelCSIToolBoxAddIn");
                System.IO.Directory.CreateDirectory(logFolder);
                string logPath = System.IO.Path.Combine(logFolder, "startup.log");
                string escapedLogPath = logPath.Replace("\"", "\"\"");
                string vbsScript = @"
Option Explicit
On Error Resume Next

Dim logPath
logPath = ""__LOG_PATH__""

Sub Log(message)
    On Error Resume Next
    Dim fso, file
    Set fso = CreateObject(""Scripting.FileSystemObject"")
    Set file = fso.OpenTextFile(logPath, 8, True)
    file.WriteLine Now & vbTab & message
    file.Close
End Sub

Function TrySetConnect(addin, targetState, actionName)
    Dim attempt
    TrySetConnect = False
    For attempt = 1 To 12
        Err.Clear
        addin.Connect = targetState
        If Err.Number = 0 Then
            WScript.Sleep 500
            Err.Clear
            If addin.Connect = targetState And Err.Number = 0 Then
                Log actionName & "" succeeded on attempt "" & attempt & "".""
                TrySetConnect = True
                Exit Function
            End If
        End If

        Log actionName & "" attempt "" & attempt & "" failed: "" & Err.Number & "" "" & Err.Description
        WScript.Sleep 700
    Next
End Function

WScript.Sleep 1200
Log ""Refresh helper started.""

Dim excel, addin, candidate, i
Set excel = Nothing
For i = 1 To 20
    Err.Clear
    Set excel = GetObject(, ""Excel.Application"")
    If Err.Number = 0 And Not excel Is Nothing Then
        Exit For
    End If
    WScript.Sleep 500
Next

If excel Is Nothing Then
    Log ""Refresh failed: could not attach to Excel.Application.""
    MsgBox ""Refresh Plugin failed: could not attach to Excel. Close and reopen Excel to load the latest build."", vbExclamation, ""Refresh Plugin""
    WScript.Quit 1
End If

Log ""Refresh helper attached to Excel.""

Set addin = Nothing
Err.Clear
Set addin = excel.COMAddIns(""ExcelCSIToolBoxAddIn"")
If Err.Number <> 0 Or addin Is Nothing Then
    Log ""Refresh helper could not get add-in by exact ProgId. Searching COMAddIns list.""
    Err.Clear
    For Each candidate In excel.COMAddIns
        If Err.Number = 0 Then
            Log ""Found COMAddIn: "" & candidate.ProgId & "", Connect="" & candidate.Connect
            If InStr(1, candidate.ProgId, ""ExcelCSIToolBoxAddIn"", vbTextCompare) > 0 Then
                Set addin = candidate
                Exit For
            End If
        End If
    Next
End If

If addin Is Nothing Then
    Log ""Refresh failed: ExcelCSIToolBoxAddIn COMAddIn was not found.""
    MsgBox ""Refresh Plugin failed: ExcelCSIToolBoxAddIn was not found in Excel COMAddIns. Close and reopen Excel to load the latest build."", vbExclamation, ""Refresh Plugin""
    WScript.Quit 1
End If

Log ""Refresh helper found add-in: "" & addin.ProgId & "", Connect="" & addin.Connect

If Not TrySetConnect(addin, False, ""Disconnect add-in"") Then
    Log ""Refresh failed: disconnect did not complete.""
    MsgBox ""Refresh Plugin failed while disconnecting the add-in. Close and reopen Excel to load the latest build."", vbExclamation, ""Refresh Plugin""
    WScript.Quit 1
End If

WScript.Sleep 1200

If Not TrySetConnect(addin, True, ""Reconnect add-in"") Then
    Log ""Refresh failed: reconnect did not complete.""
    MsgBox ""Refresh Plugin failed while reconnecting the add-in. Close and reopen Excel to load the latest build."", vbExclamation, ""Refresh Plugin""
    WScript.Quit 1
End If

Log ""Refresh helper completed.""
".Replace("__LOG_PATH__", escapedLogPath);
                System.IO.File.WriteAllText(tempVbsPath, vbsScript, System.Text.Encoding.ASCII);
                System.Diagnostics.Process.Start("wscript.exe", "\"" + tempVbsPath + "\"");
                AddInDiagnostics.Log("Refresh helper launched: " + tempVbsPath);
            }
            catch (Exception ex)
            {
                AddInDiagnostics.LogException("Refresh Plugin", ex);
                System.Windows.Forms.MessageBox.Show(
                    "Failed to restart the add-in: " + ex.Message,
                    "Refresh Plugin Error",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        private static void CloseModelessWpfWindowsForRefresh()
        {
            try
            {
                var application = System.Windows.Application.Current;
                if (application == null)
                {
                    return;
                }

                var windows = new System.Collections.Generic.List<System.Windows.Window>();
                foreach (System.Windows.Window window in application.Windows)
                {
                    windows.Add(window);
                }

                foreach (System.Windows.Window window in windows)
                {
                    window.Close();
                }

                AddInDiagnostics.Log("Closed " + windows.Count + " WPF window(s) before refresh.");
            }
            catch (Exception ex)
            {
                AddInDiagnostics.Log("Could not close WPF windows before refresh: " + ex.Message);
            }
        }
    }
}
