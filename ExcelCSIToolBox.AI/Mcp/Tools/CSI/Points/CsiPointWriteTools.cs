using ExcelCSIToolBox.AI.Mcp.Contracts;
using ExcelCSIToolBox.AI.Mcp.Tools.CSI.Base;
using ExcelCSIToolBox.Core.Abstractions.CSI;
using ExcelCSIToolBox.Core.Common.Results;
using ExcelCSIToolBox.Core.Models.CSI;
using ExcelCSIToolBox.Infrastructure.CSISapModel;
using System.Collections.Generic;

namespace ExcelCSIToolBox.AI.Mcp.Tools.CSI.Points
{
    public sealed class PointsAddByCoordinatesTool : CsiWriteToolBase<PointByCoordinatesArgs>
    {
        public PointsAddByCoordinatesTool(ICsiModelCommandService commandService) : base(commandService) { }
        public override string Name => "points.add_by_coordinates";
        public override string Title => "Add Point By Coordinates";
        public override string Category => "Points";
        public override string SubCategory => "Creation";
        public override string Description => "Adds one CSI point by Cartesian coordinates. Low-risk write tool with dry-run support.";
        public override CsiMethodRiskLevel RiskLevel => CsiMethodRiskLevel.Low;
        public override bool RequiresConfirmation => false;
        public override bool SupportsDryRun => true;

        protected override ToolCallResponse Execute(PointByCoordinatesArgs args)
        {
            string userName = args.GetUserName();
            return args.DryRun
                ? Preview(CommandService.PreviewAddPoint(args.X, args.Y, args.Z, userName))
                : Result(CommandService.AddPoint(args.X, args.Y, args.Z, userName, args.Confirmed));
        }
    }

    public sealed class PointsAddCartesianTool : CsiWriteToolBase<PointByCoordinatesArgs>
    {
        public PointsAddCartesianTool(ICsiModelCommandService commandService) : base(commandService) { }
        public override string Name => "points.add_cartesian";
        public override string Title => "Add Point Cartesian";
        public override string Category => "Points";
        public override string SubCategory => "Creation";
        public override string Description => "Adds one CSI point by Cartesian coordinates. Alias for points.add_by_coordinates.";
        public override CsiMethodRiskLevel RiskLevel => CsiMethodRiskLevel.Low;
        public override bool RequiresConfirmation => false;
        public override bool SupportsDryRun => true;

        protected override ToolCallResponse Execute(PointByCoordinatesArgs args)
        {
            string userName = args.GetUserName();
            return args.DryRun
                ? Preview(CommandService.PreviewAddPoint(args.X, args.Y, args.Z, userName))
                : Result(CommandService.AddPoint(args.X, args.Y, args.Z, userName, args.Confirmed));
        }
    }

    public sealed class PointsSetRestraintTool : CsiActiveConnectionToolBase
    {
        private readonly CsiWriteGuard _writeGuard = new CsiWriteGuard();
        private readonly CsiOperationLogger _logger = new CsiOperationLogger();

        public PointsSetRestraintTool(ICSISapModelConnectionService etabsService, ICSISapModelConnectionService sap2000Service) : base(etabsService, sap2000Service) { }
        public override string Name => "points.set_restraint";
        public override string Title => "Set Point Restraint";
        public override string Category => "Points";
        public override string SubCategory => "Assignments";
        public override string Description => "Assigns restraints to point objects. Medium risk and requires confirmation.";
        public override bool IsReadOnly => false;
        public override CsiMethodRiskLevel RiskLevel => CsiMethodRiskLevel.Medium;
        public override bool RequiresConfirmation => true;
        public override bool SupportsDryRun => true;

        protected override ToolCallResponse Execute(ICSISapModelConnectionService service, string argumentsJson)
        {
            PointRestraintArgs args = ReadArgs<PointRestraintArgs>(argumentsJson);
            IReadOnlyList<string> names = args.PointNames ?? new List<string>();
            if (args.DryRun)
            {
                return Preview(new CsiWritePreview
                {
                    OperationName = Name,
                    RiskLevel = RiskLevel,
                    RequiresConfirmation = true,
                    SupportsDryRun = true,
                    Summary = $"This will set restraints for {Count(names)} point object(s).",
                    AffectedObjects = names
                });
            }

            OperationResult guardResult = _writeGuard.ValidateWrite(Name, RiskLevel, args.Confirmed, names);
            if (!guardResult.IsSuccess)
            {
                _logger.Log(service.ProductName, Name, Category, SubCategory, RiskLevel, $"count={Count(names)}", names, args.Confirmed, false, guardResult.Message);
                return Result(guardResult);
            }

            OperationResult result = service.SetPointRestraint(names, args.Restraints);
            _logger.Log(service.ProductName, Name, Category, SubCategory, RiskLevel, $"count={Count(names)}", names, args.Confirmed, result.IsSuccess, result.Message);
            return Result(result);
        }
    }

    public sealed class PointsSetLoadForceTool : CsiActiveConnectionToolBase
    {
        private readonly CsiWriteGuard _writeGuard = new CsiWriteGuard();
        private readonly CsiOperationLogger _logger = new CsiOperationLogger();

        public PointsSetLoadForceTool(ICSISapModelConnectionService etabsService, ICSISapModelConnectionService sap2000Service) : base(etabsService, sap2000Service) { }
        public override string Name => "points.set_load_force";
        public override string Title => "Set Point Load Force";
        public override string Category => "Points";
        public override string SubCategory => "Loads";
        public override string Description => "Assigns point force loads to point objects. Medium risk and requires confirmation.";
        public override bool IsReadOnly => false;
        public override CsiMethodRiskLevel RiskLevel => CsiMethodRiskLevel.Medium;
        public override bool RequiresConfirmation => true;
        public override bool SupportsDryRun => true;

        protected override ToolCallResponse Execute(ICSISapModelConnectionService service, string argumentsJson)
        {
            PointLoadForceArgs args = ReadArgs<PointLoadForceArgs>(argumentsJson);
            IReadOnlyList<string> names = args.PointNames ?? new List<string>();
            if (args.DryRun)
            {
                return Preview(new CsiWritePreview
                {
                    OperationName = Name,
                    RiskLevel = RiskLevel,
                    RequiresConfirmation = true,
                    SupportsDryRun = true,
                    Summary = $"This will assign point load pattern '{args.LoadPattern}' to {Count(names)} point object(s).",
                    AffectedObjects = names
                });
            }

            OperationResult guardResult = _writeGuard.ValidateWrite(Name, RiskLevel, args.Confirmed, names);
            if (!guardResult.IsSuccess)
            {
                _logger.Log(service.ProductName, Name, Category, SubCategory, RiskLevel, $"count={Count(names)} loadPattern={args.LoadPattern}", names, args.Confirmed, false, guardResult.Message);
                return Result(guardResult);
            }

            OperationResult result = service.SetPointLoadForce(names, args.LoadPattern, args.ForceValues, args.Replace, args.CoordinateSystem);
            _logger.Log(service.ProductName, Name, Category, SubCategory, RiskLevel, $"count={Count(names)} loadPattern={args.LoadPattern}", names, args.Confirmed, result.IsSuccess, result.Message);
            return Result(result);
        }
    }

    public sealed class PointByCoordinatesArgs : LowRiskWriteArgs
    {
        public double X { get; set; }
        public double Y { get; set; }
        public double Z { get; set; }
        public string UserName { get; set; }
        public string Name { get; set; }
        public string PointName { get; set; }
        public string UniqueName { get; set; }

        public string GetUserName()
        {
            if (!string.IsNullOrWhiteSpace(UserName))
            {
                return UserName;
            }

            if (!string.IsNullOrWhiteSpace(UniqueName))
            {
                return UniqueName;
            }

            if (!string.IsNullOrWhiteSpace(PointName))
            {
                return PointName;
            }

            return Name;
        }
    }

    public sealed class PointRestraintArgs : DryRunConfirmedArgs
    {
        public List<string> PointNames { get; set; }
        public List<bool> Restraints { get; set; }
    }

    public sealed class PointLoadForceArgs : DryRunConfirmedArgs
    {
        public List<string> PointNames { get; set; }
        public string LoadPattern { get; set; }
        public List<double> ForceValues { get; set; }
        public bool Replace { get; set; } = true;
        public string CoordinateSystem { get; set; } = "Global";
    }
}
