using System.Collections.Generic;
using ExcelCSIToolBox.AI.Mcp.Contracts;
using ExcelCSIToolBox.AI.Mcp.Tools.CSI.Base;
using ExcelCSIToolBox.Core.Abstractions.CSI;
using ExcelCSIToolBox.Core.Common.Results;
using ExcelCSIToolBox.Core.Models.CSI;
using ExcelCSIToolBox.Infrastructure.CSISapModel;

namespace ExcelCSIToolBox.AI.Mcp.Tools.CSI.Loads.LoadCombinations
{
    public sealed class LoadCombinationsGetAllTool : CsiActiveConnectionToolBase
    {
        public LoadCombinationsGetAllTool(ICSISapModelConnectionService etabsService, ICSISapModelConnectionService sap2000Service) : base(etabsService, sap2000Service) { }
        public override string Name => "loads.combinations.get_all";
        public override string Title => "Get Load Combinations";
        public override string Category => "Loads";
        public override string SubCategory => "Combinations";
        public override string Description => "Returns load combination names and types from the active CSI model.";
        public override bool IsReadOnly => true;
        public override CsiMethodRiskLevel RiskLevel => CsiMethodRiskLevel.None;
        public override bool RequiresConfirmation => false;
        public override bool SupportsDryRun => false;

        protected override ToolCallResponse Execute(ICSISapModelConnectionService service, string argumentsJson)
        {
            return Result(service.GetLoadCombinations());
        }
    }

    public sealed class LoadCombinationsGetDetailsTool : CsiActiveConnectionToolBase
    {
        public LoadCombinationsGetDetailsTool(ICSISapModelConnectionService etabsService, ICSISapModelConnectionService sap2000Service) : base(etabsService, sap2000Service) { }
        public override string Name => "loads.combinations.get_details";
        public override string Title => "Get Load Combination Details";
        public override string Category => "Loads";
        public override string SubCategory => "Combinations";
        public override string Description => "Returns load cases and scale factors for one load combination.";
        public override bool IsReadOnly => true;
        public override CsiMethodRiskLevel RiskLevel => CsiMethodRiskLevel.None;
        public override bool RequiresConfirmation => false;
        public override bool SupportsDryRun => false;

        protected override ToolCallResponse Execute(ICSISapModelConnectionService service, string argumentsJson)
        {
            CombinationNameArgs args = ReadArgs<CombinationNameArgs>(argumentsJson);
            return Result(service.GetLoadCombinationDetails(args.CombinationName));
        }
    }

    public sealed class LoadCombinationsDeleteTool : CsiActiveConnectionToolBase
    {
        private readonly CsiWriteGuard _writeGuard = new CsiWriteGuard();
        private readonly CsiOperationLogger _logger = new CsiOperationLogger();

        public LoadCombinationsDeleteTool(ICSISapModelConnectionService etabsService, ICSISapModelConnectionService sap2000Service) : base(etabsService, sap2000Service) { }
        public override string Name => "loads.combinations.delete";
        public override string Title => "Delete Load Combinations";
        public override string Category => "Loads";
        public override string SubCategory => "Combinations";
        public override string Description => "Deletes load combinations. High risk and requires explicit confirmation.";
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
                    Summary = $"This will delete {Count(names)} load combination(s). This is high risk.",
                    AffectedObjects = names
                });
            }

            OperationResult guardResult = _writeGuard.ValidateWrite(Name, RiskLevel, args.Confirmed, names);
            if (!guardResult.IsSuccess)
            {
                _logger.Log(service.ProductName, Name, Category, SubCategory, RiskLevel, $"count={Count(names)}", names, args.Confirmed, false, guardResult.Message);
                return Result(guardResult);
            }

            OperationResult result = service.DeleteLoadCombinations(names);
            _logger.Log(service.ProductName, Name, Category, SubCategory, RiskLevel, $"count={Count(names)}", names, args.Confirmed, result.IsSuccess, result.Message);
            return Result(result);
        }
    }

    public sealed class CombinationNameArgs
    {
        public string CombinationName { get; set; }
    }
}
