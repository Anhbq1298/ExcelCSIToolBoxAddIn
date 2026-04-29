using System.Collections.Generic;
using ExcelCSIToolBox.AI.Mcp.Contracts;
using ExcelCSIToolBox.AI.Mcp.Tools.CSI.Base;
using ExcelCSIToolBox.Core.Abstractions.CSI;
using ExcelCSIToolBox.Core.Common.Results;
using ExcelCSIToolBox.Core.Models.CSI;
using ExcelCSIToolBox.Infrastructure.CSISapModel;

namespace ExcelCSIToolBox.AI.Mcp.Tools.CSI.Loads.LoadPatterns
{
    public sealed class LoadPatternsGetAllTool : CsiActiveConnectionToolBase
    {
        public LoadPatternsGetAllTool(ICSISapModelConnectionService etabsService, ICSISapModelConnectionService sap2000Service) : base(etabsService, sap2000Service) { }
        public override string Name => "loads.patterns.get_all";
        public override string Title => "Get Load Patterns";
        public override string Category => "Loads";
        public override string SubCategory => "Load Patterns";
        public override string Description => "Returns load pattern names and types from the active CSI model.";
        public override bool IsReadOnly => true;
        public override CsiMethodRiskLevel RiskLevel => CsiMethodRiskLevel.None;
        public override bool RequiresConfirmation => false;
        public override bool SupportsDryRun => false;

        protected override ToolCallResponse Execute(ICSISapModelConnectionService service, string argumentsJson)
        {
            return Result(service.GetLoadPatterns());
        }
    }

    public sealed class LoadPatternsDeleteTool : CsiActiveConnectionToolBase
    {
        private readonly CsiWriteGuard _writeGuard = new CsiWriteGuard();
        private readonly CsiOperationLogger _logger = new CsiOperationLogger();

        public LoadPatternsDeleteTool(ICSISapModelConnectionService etabsService, ICSISapModelConnectionService sap2000Service) : base(etabsService, sap2000Service) { }
        public override string Name => "loads.patterns.delete";
        public override string Title => "Delete Load Patterns";
        public override string Category => "Loads";
        public override string SubCategory => "Load Patterns";
        public override string Description => "Deletes load patterns. High risk and requires explicit confirmation.";
        public override bool IsReadOnly => false;
        public override CsiMethodRiskLevel RiskLevel => CsiMethodRiskLevel.High;
        public override bool RequiresConfirmation => true;
        public override bool SupportsDryRun => true;

        protected override ToolCallResponse Execute(ICSISapModelConnectionService service, string argumentsJson)
        {
            NamesDryRunArgs args = ReadArgs<NamesDryRunArgs>(argumentsJson);
            IReadOnlyList<string> names = args.Names ?? new List<string>();
            if (args.DryRun)
            {
                return Preview(new CsiWritePreview
                {
                    OperationName = Name,
                    RiskLevel = RiskLevel,
                    RequiresConfirmation = true,
                    SupportsDryRun = true,
                    Summary = $"This will delete {Count(names)} load pattern(s). This is high risk.",
                    AffectedObjects = names
                });
            }

            OperationResult guardResult = _writeGuard.ValidateWrite(Name, RiskLevel, args.Confirmed, names);
            if (!guardResult.IsSuccess)
            {
                _logger.Log(service.ProductName, Name, Category, SubCategory, RiskLevel, $"count={Count(names)}", names, args.Confirmed, false, guardResult.Message);
                return Result(guardResult);
            }

            OperationResult result = service.DeleteLoadPatterns(names);
            _logger.Log(service.ProductName, Name, Category, SubCategory, RiskLevel, $"count={Count(names)}", names, args.Confirmed, result.IsSuccess, result.Message);
            return Result(result);
        }
    }
}
