using System.Collections.Generic;
using ExcelCSIToolBox.AI.Mcp.Contracts;
using ExcelCSIToolBox.AI.Mcp.Tools.CSI.Base;
using ExcelCSIToolBox.Core.Abstractions.CSI;
using ExcelCSIToolBox.Core.Models.CSI;

namespace ExcelCSIToolBox.AI.Mcp.Tools.CSI.Points
{
    public sealed class PointsGetSelectedTool : CsiActiveConnectionToolBase
    {
        public PointsGetSelectedTool(ICSISapModelConnectionService etabsService, ICSISapModelConnectionService sap2000Service) : base(etabsService, sap2000Service) { }
        public override string Name => "points.get_selected";
        public override string Title => "Get Selected Points";
        public override string Category => "Points";
        public override string SubCategory => "Read";
        public override string Description => "Returns selected point objects with coordinates from the active CSI model.";
        public override bool IsReadOnly => true;
        public override CsiMethodRiskLevel RiskLevel => CsiMethodRiskLevel.None;
        public override bool RequiresConfirmation => false;
        public override bool SupportsDryRun => false;

        protected override ToolCallResponse Execute(ICSISapModelConnectionService service, string argumentsJson)
        {
            return Result(service.GetSelectedPointsFromActiveModel());
        }
    }

    public sealed class PointsGetAllNamesTool : CsiActiveConnectionToolBase
    {
        public PointsGetAllNamesTool(ICSISapModelConnectionService etabsService, ICSISapModelConnectionService sap2000Service) : base(etabsService, sap2000Service) { }
        public override string Name => "points.get_all_names";
        public override string Title => "Get Point Names";
        public override string Category => "Points";
        public override string SubCategory => "Read";
        public override string Description => "Returns all point object names from the active CSI model.";
        public override bool IsReadOnly => true;
        public override CsiMethodRiskLevel RiskLevel => CsiMethodRiskLevel.None;
        public override bool RequiresConfirmation => false;
        public override bool SupportsDryRun => false;

        protected override ToolCallResponse Execute(ICSISapModelConnectionService service, string argumentsJson)
        {
            return Result(service.GetPointNames());
        }
    }

    public sealed class PointsGetByNameTool : CsiActiveConnectionToolBase
    {
        public PointsGetByNameTool(ICSISapModelConnectionService etabsService, ICSISapModelConnectionService sap2000Service) : base(etabsService, sap2000Service) { }
        public override string Name => "points.get_by_name";
        public override string Title => "Get Point By Name";
        public override string Category => "Points";
        public override string SubCategory => "Read";
        public override string Description => "Returns point object coordinates and selection state by name.";
        public override bool IsReadOnly => true;
        public override CsiMethodRiskLevel RiskLevel => CsiMethodRiskLevel.None;
        public override bool RequiresConfirmation => false;
        public override bool SupportsDryRun => false;

        protected override ToolCallResponse Execute(ICSISapModelConnectionService service, string argumentsJson)
        {
            PointNameArgs args = ReadArgs<PointNameArgs>(argumentsJson);
            return Result(service.GetPointByName(args.PointName));
        }
    }

    public sealed class PointsGetCoordinatesTool : CsiActiveConnectionToolBase
    {
        public PointsGetCoordinatesTool(ICSISapModelConnectionService etabsService, ICSISapModelConnectionService sap2000Service) : base(etabsService, sap2000Service) { }
        public override string Name => "points.get_coordinates";
        public override string Title => "Get Point Coordinates";
        public override string Category => "Points";
        public override string SubCategory => "Read";
        public override string Description => "Returns Cartesian coordinates for one point object.";
        public override bool IsReadOnly => true;
        public override CsiMethodRiskLevel RiskLevel => CsiMethodRiskLevel.None;
        public override bool RequiresConfirmation => false;
        public override bool SupportsDryRun => false;

        protected override ToolCallResponse Execute(ICSISapModelConnectionService service, string argumentsJson)
        {
            PointNameArgs args = ReadArgs<PointNameArgs>(argumentsJson);
            return Result(service.GetPointCoordinates(args.PointName));
        }
    }

    public sealed class PointsGetRestraintTool : CsiActiveConnectionToolBase
    {
        public PointsGetRestraintTool(ICSISapModelConnectionService etabsService, ICSISapModelConnectionService sap2000Service) : base(etabsService, sap2000Service) { }
        public override string Name => "points.get_restraint";
        public override string Title => "Get Point Restraint";
        public override string Category => "Points";
        public override string SubCategory => "Read";
        public override string Description => "Returns restraint assignments for one point object.";
        public override bool IsReadOnly => true;
        public override CsiMethodRiskLevel RiskLevel => CsiMethodRiskLevel.None;
        public override bool RequiresConfirmation => false;
        public override bool SupportsDryRun => false;

        protected override ToolCallResponse Execute(ICSISapModelConnectionService service, string argumentsJson)
        {
            PointNameArgs args = ReadArgs<PointNameArgs>(argumentsJson);
            return Result(service.GetPointRestraint(args.PointName));
        }
    }

    public sealed class PointsGetLoadForcesTool : CsiActiveConnectionToolBase
    {
        public PointsGetLoadForcesTool(ICSISapModelConnectionService etabsService, ICSISapModelConnectionService sap2000Service) : base(etabsService, sap2000Service) { }
        public override string Name => "points.get_load_forces";
        public override string Title => "Get Point Load Forces";
        public override string Category => "Points";
        public override string SubCategory => "Loads";
        public override string Description => "Returns point force load assignments for one point object.";
        public override bool IsReadOnly => true;
        public override CsiMethodRiskLevel RiskLevel => CsiMethodRiskLevel.None;
        public override bool RequiresConfirmation => false;
        public override bool SupportsDryRun => false;

        protected override ToolCallResponse Execute(ICSISapModelConnectionService service, string argumentsJson)
        {
            PointNameArgs args = ReadArgs<PointNameArgs>(argumentsJson);
            return Result(service.GetPointLoadForces(args.PointName));
        }
    }

    public sealed class PointsSetSelectedTool : CsiActiveConnectionToolBase
    {
        public PointsSetSelectedTool(ICSISapModelConnectionService etabsService, ICSISapModelConnectionService sap2000Service) : base(etabsService, sap2000Service) { }
        public override string Name => "points.set_selected";
        public override string Title => "Select Points By Name";
        public override string Category => "Points";
        public override string SubCategory => "Selection";
        public override string Description => "Selects point objects by unique names. Low-risk selection write tool with dry-run support.";
        public override bool IsReadOnly => false;
        public override CsiMethodRiskLevel RiskLevel => CsiMethodRiskLevel.Low;
        public override bool RequiresConfirmation => false;
        public override bool SupportsDryRun => true;

        protected override ToolCallResponse Execute(ICSISapModelConnectionService service, string argumentsJson)
        {
            NamesDryRunArgs args = ReadArgs<NamesDryRunArgs>(argumentsJson);
            if (args.DryRun)
            {
                return Preview(new CsiWritePreview
                {
                    OperationName = Name,
                    RiskLevel = RiskLevel,
                    RequiresConfirmation = false,
                    SupportsDryRun = true,
                    Summary = $"This will select {Count(args.Names)} point object(s).",
                    AffectedObjects = args.Names ?? new List<string>()
                });
            }

            return Result(service.SelectPointsByUniqueNames(args.Names));
        }
    }

    public sealed class PointNameArgs
    {
        public string PointName { get; set; }
    }
}
