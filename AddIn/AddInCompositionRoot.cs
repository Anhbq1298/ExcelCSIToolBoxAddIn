using System;
using ExcelCSIToolBox.AI.Agent;
using ExcelCSIToolBox.AI.Mcp.Client;
using ExcelCSIToolBox.AI.Mcp.Server;
using ExcelCSIToolBox.AI.Ollama;
using ExcelCSIToolBox.Core.Abstractions.CSI;
using ExcelCSIToolBox.Core.Abstractions.Excel;
using ExcelCSIToolBox.Infrastructure.CSISapModel;
using ExcelCSIToolBox.Infrastructure.CSISapModel.Random;
using ExcelCSIToolBox.Infrastructure.CSISapModel.ReadOnly;
using ExcelCSIToolBox.Infrastructure.CSISapModel.Truss;
using ExcelCSIToolBox.Infrastructure.CSISapModel.Workflow;
using ExcelCSIToolBox.Infrastructure.Excel;
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

        public static void Configure(
            ICSISapModelConnectionService etabsConnectionService,
            ICSISapModelConnectionService sap2000ConnectionService)
        {
            _etabsConnectionService = etabsConnectionService ?? throw new ArgumentNullException(nameof(etabsConnectionService));
            _sap2000ConnectionService = sap2000ConnectionService ?? throw new ArgumentNullException(nameof(sap2000ConnectionService));
            _excelSelectionService = new ExcelSelectionService();
            _excelOutputService = new ExcelOutputService();

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
            return new AiAgentChatControl(new AiAgentChatViewModel(sessionService));
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
                new CsiWorkflowExecutionService());

            var mcpServer = new LocalMcpServer(context);
            var mcpClient = new LocalMcpClient(mcpServer);
            var ollamaService = new OllamaChatService();
            var orchestrator = new AiAgentOrchestrator(ollamaService, mcpClient);

            return new AiChatSessionService(
                orchestrator,
                _etabsConnectionService,
                _sap2000ConnectionService,
                OllamaChatService.DefaultModel);
        }

        private static void EnsureConfigured()
        {
            if (_etabsConnectionService == null || _sap2000ConnectionService == null)
            {
                throw new InvalidOperationException("The add-in composition root is not configured.");
            }
        }
    }
}
