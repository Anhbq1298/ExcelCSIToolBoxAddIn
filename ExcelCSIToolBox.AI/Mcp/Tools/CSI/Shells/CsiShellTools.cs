using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using ExcelCSIToolBox.AI.Mcp.Contracts;
using ExcelCSIToolBox.AI.Mcp.Tools.CSI.Base;
using ExcelCSIToolBox.Core.Abstractions.CSI;
using ExcelCSIToolBox.Core.Common.Results;
using ExcelCSIToolBox.Core.Models.CSI;
using ExcelCSIToolBox.Data.DTOs.CSI;
using ExcelCSIToolBox.Data.Models;
using Newtonsoft.Json;

namespace ExcelCSIToolBox.AI.Mcp.Tools.CSI.Shells
{
    public abstract class CsiShellToolBase : IMcpTool, IMcpToolMetadata
    {
        private readonly ICSISapModelConnectionService _etabsService;
        private readonly ICSISapModelConnectionService _sap2000Service;

        protected CsiShellToolBase(
            ICSISapModelConnectionService etabsService,
            ICSISapModelConnectionService sap2000Service)
        {
            _etabsService = etabsService ?? throw new ArgumentNullException(nameof(etabsService));
            _sap2000Service = sap2000Service ?? throw new ArgumentNullException(nameof(sap2000Service));
        }

        public abstract string Name { get; }
        public abstract string Title { get; }
        public string Category => "Shells / Areas";
        public abstract string SubCategory { get; }
        public abstract string Description { get; }
        public abstract bool IsReadOnly { get; }
        public abstract CsiMethodRiskLevel RiskLevel { get; }
        public abstract bool RequiresConfirmation { get; }
        public abstract bool SupportsDryRun { get; }

        public Task<ToolCallResponse> ExecuteAsync(string argumentsJson, CancellationToken cancellationToken)
        {
            try
            {
                OperationResult<ICSISapModelConnectionService> serviceResult = GetActiveService();
                if (!serviceResult.IsSuccess)
                {
                    return Task.FromResult(Fail(serviceResult.Message));
                }

                return Task.FromResult(Execute(serviceResult.Data, argumentsJson ?? "{}"));
            }
            catch (Exception ex)
            {
                return Task.FromResult(Fail("Shell/area tool execution failed: " + ex.Message));
            }
        }

        protected abstract ToolCallResponse Execute(ICSISapModelConnectionService service, string argumentsJson);

        protected TArgs ReadArgs<TArgs>(string argumentsJson) where TArgs : class, new()
        {
            return JsonConvert.DeserializeObject<TArgs>(argumentsJson ?? "{}") ?? new TArgs();
        }

        protected ToolCallResponse Result<T>(OperationResult<T> result)
        {
            if (!result.IsSuccess)
            {
                return Fail(result.Message);
            }

            return Ok(result.Message ?? "Shell/area query completed.", result.Data);
        }

        protected ToolCallResponse Result(OperationResult result)
        {
            return new ToolCallResponse
            {
                ToolName = Name,
                Success = result.IsSuccess,
                Message = result.Message,
                ResultJson = JsonConvert.SerializeObject(new
                {
                    result.IsSuccess,
                    result.Message
                })
            };
        }

        protected ToolCallResponse Preview(CsiWritePreview preview)
        {
            return Ok(preview.Summary, preview);
        }

        protected ToolCallResponse Ok(string message, object payload)
        {
            return new ToolCallResponse
            {
                ToolName = Name,
                Success = true,
                Message = message,
                ResultJson = JsonConvert.SerializeObject(payload)
            };
        }

        protected ToolCallResponse Fail(string message)
        {
            return new ToolCallResponse
            {
                ToolName = Name,
                Success = false,
                Message = message,
                ResultJson = null
            };
        }

        private OperationResult<ICSISapModelConnectionService> GetActiveService()
        {
            OperationResult<CSISapModelConnectionInfoDTO> etabs = _etabsService.GetCurrentConnection();
            if (IsConnected(etabs))
            {
                return OperationResult<ICSISapModelConnectionService>.Success(_etabsService);
            }

            OperationResult<CSISapModelConnectionInfoDTO> sap2000 = _sap2000Service.GetCurrentConnection();
            if (IsConnected(sap2000))
            {
                return OperationResult<ICSISapModelConnectionService>.Success(_sap2000Service);
            }

            etabs = _etabsService.TryAttachToRunningInstance();
            if (IsConnected(etabs))
            {
                return OperationResult<ICSISapModelConnectionService>.Success(_etabsService);
            }

            sap2000 = _sap2000Service.TryAttachToRunningInstance();
            if (IsConnected(sap2000))
            {
                return OperationResult<ICSISapModelConnectionService>.Success(_sap2000Service);
            }

            return OperationResult<ICSISapModelConnectionService>.Failure("No ETABS or SAP2000 model is attached.");
        }

        private static bool IsConnected(OperationResult<CSISapModelConnectionInfoDTO> result)
        {
            return result != null &&
                   result.IsSuccess &&
                   result.Data != null &&
                   result.Data.IsConnected &&
                   result.Data.SapModel != null;
        }
    }

    public sealed class ShellsGetByNameTool : CsiShellToolBase
    {
        public ShellsGetByNameTool(ICSISapModelConnectionService etabsService, ICSISapModelConnectionService sap2000Service) : base(etabsService, sap2000Service) { }
        public override string Name => "shells.get_by_name";
        public override string Title => "Get Shell / Area By Name";
        public override string SubCategory => "Read";
        public override string Description => "Returns shell/area points, property, and selection status for one area object.";
        public override bool IsReadOnly => true;
        public override CsiMethodRiskLevel RiskLevel => CsiMethodRiskLevel.None;
        public override bool RequiresConfirmation => false;
        public override bool SupportsDryRun => false;

        protected override ToolCallResponse Execute(ICSISapModelConnectionService service, string argumentsJson)
        {
            ShellNameArgs args = ReadArgs<ShellNameArgs>(argumentsJson);
            return Result(service.GetShellByName(args.AreaName));
        }
    }

    public sealed class ShellsGetPointsTool : CsiShellToolBase
    {
        public ShellsGetPointsTool(ICSISapModelConnectionService etabsService, ICSISapModelConnectionService sap2000Service) : base(etabsService, sap2000Service) { }
        public override string Name => "shells.get_points";
        public override string Title => "Get Shell / Area Points";
        public override string SubCategory => "Read";
        public override string Description => "Returns point names that define one shell/area object.";
        public override bool IsReadOnly => true;
        public override CsiMethodRiskLevel RiskLevel => CsiMethodRiskLevel.None;
        public override bool RequiresConfirmation => false;
        public override bool SupportsDryRun => false;

        protected override ToolCallResponse Execute(ICSISapModelConnectionService service, string argumentsJson)
        {
            ShellNameArgs args = ReadArgs<ShellNameArgs>(argumentsJson);
            return Result(service.GetShellPoints(args.AreaName));
        }
    }

    public sealed class ShellsGetPropertyTool : CsiShellToolBase
    {
        public ShellsGetPropertyTool(ICSISapModelConnectionService etabsService, ICSISapModelConnectionService sap2000Service) : base(etabsService, sap2000Service) { }
        public override string Name => "shells.get_property";
        public override string Title => "Get Shell / Area Property";
        public override string SubCategory => "Read";
        public override string Description => "Returns the assigned shell/area property name.";
        public override bool IsReadOnly => true;
        public override CsiMethodRiskLevel RiskLevel => CsiMethodRiskLevel.None;
        public override bool RequiresConfirmation => false;
        public override bool SupportsDryRun => false;

        protected override ToolCallResponse Execute(ICSISapModelConnectionService service, string argumentsJson)
        {
            ShellNameArgs args = ReadArgs<ShellNameArgs>(argumentsJson);
            return Result(service.GetShellProperty(args.AreaName));
        }
    }

    public sealed class ShellsGetSelectedTool : CsiShellToolBase
    {
        public ShellsGetSelectedTool(ICSISapModelConnectionService etabsService, ICSISapModelConnectionService sap2000Service) : base(etabsService, sap2000Service) { }
        public override string Name => "shells.get_selected";
        public override string Title => "Get Selected Shell / Area Objects";
        public override string SubCategory => "Read";
        public override string Description => "Returns selected shell/area object names from the active CSI model.";
        public override bool IsReadOnly => true;
        public override CsiMethodRiskLevel RiskLevel => CsiMethodRiskLevel.None;
        public override bool RequiresConfirmation => false;
        public override bool SupportsDryRun => false;

        protected override ToolCallResponse Execute(ICSISapModelConnectionService service, string argumentsJson)
        {
            return Result(service.GetSelectedShells());
        }
    }

    public sealed class ShellsGetUniformLoadsTool : CsiShellToolBase
    {
        public ShellsGetUniformLoadsTool(ICSISapModelConnectionService etabsService, ICSISapModelConnectionService sap2000Service) : base(etabsService, sap2000Service) { }
        public override string Name => "shells.get_uniform_loads";
        public override string Title => "Get Shell / Area Uniform Loads";
        public override string SubCategory => "Loads";
        public override string Description => "Returns uniform load assignments for one shell/area object.";
        public override bool IsReadOnly => true;
        public override CsiMethodRiskLevel RiskLevel => CsiMethodRiskLevel.None;
        public override bool RequiresConfirmation => false;
        public override bool SupportsDryRun => false;

        protected override ToolCallResponse Execute(ICSISapModelConnectionService service, string argumentsJson)
        {
            ShellNameArgs args = ReadArgs<ShellNameArgs>(argumentsJson);
            return Result(service.GetShellUniformLoads(args.AreaName));
        }
    }

    public sealed class ShellsAddByPointsTool : CsiShellToolBase
    {
        public ShellsAddByPointsTool(ICSISapModelConnectionService etabsService, ICSISapModelConnectionService sap2000Service) : base(etabsService, sap2000Service) { }
        public override string Name => "shells.add_by_points";
        public override string Title => "Add Shell / Area By Points";
        public override string SubCategory => "Creation";
        public override string Description => "Adds one shell/area object by existing point names. Low-risk write tool with dry-run support.";
        public override bool IsReadOnly => false;
        public override CsiMethodRiskLevel RiskLevel => CsiMethodRiskLevel.Low;
        public override bool RequiresConfirmation => false;
        public override bool SupportsDryRun => true;

        protected override ToolCallResponse Execute(ICSISapModelConnectionService service, string argumentsJson)
        {
            ShellByPointsArgs args = ReadArgs<ShellByPointsArgs>(argumentsJson);
            return args.DryRun
                ? Preview(service.PreviewAddShellByPoint(args.PointNames, args.PropertyName, args.UserName))
                : Result(service.AddShellByPoint(args.PointNames, args.PropertyName, args.UserName, args.Confirmed));
        }
    }

    public sealed class ShellsAddByCoordinatesTool : CsiShellToolBase
    {
        public ShellsAddByCoordinatesTool(ICSISapModelConnectionService etabsService, ICSISapModelConnectionService sap2000Service) : base(etabsService, sap2000Service) { }
        public override string Name => "shells.add_by_coordinates";
        public override string Title => "Add Shell / Area By Coordinates";
        public override string SubCategory => "Creation";
        public override string Description => "Adds one shell/area object by coordinates. Low-risk write tool with dry-run support.";
        public override bool IsReadOnly => false;
        public override CsiMethodRiskLevel RiskLevel => CsiMethodRiskLevel.Low;
        public override bool RequiresConfirmation => false;
        public override bool SupportsDryRun => true;

        protected override ToolCallResponse Execute(ICSISapModelConnectionService service, string argumentsJson)
        {
            ShellByCoordinatesArgs args = ReadArgs<ShellByCoordinatesArgs>(argumentsJson);
            return args.DryRun
                ? Preview(service.PreviewAddShellByCoord(args.Points, args.PropertyName, args.UserName, args.CoordinateSystem))
                : Result(service.AddShellByCoord(args.Points, args.PropertyName, args.UserName, args.CoordinateSystem, args.Confirmed));
        }
    }

    public sealed class ShellsAssignUniformLoadTool : CsiShellToolBase
    {
        public ShellsAssignUniformLoadTool(ICSISapModelConnectionService etabsService, ICSISapModelConnectionService sap2000Service) : base(etabsService, sap2000Service) { }
        public override string Name => "shells.assign_uniform_load";
        public override string Title => "Assign Shell / Area Uniform Load";
        public override string SubCategory => "Loads";
        public override string Description => "Assigns uniform load to shell/area objects. Medium risk and requires confirmation.";
        public override bool IsReadOnly => false;
        public override CsiMethodRiskLevel RiskLevel => CsiMethodRiskLevel.Medium;
        public override bool RequiresConfirmation => true;
        public override bool SupportsDryRun => true;

        protected override ToolCallResponse Execute(ICSISapModelConnectionService service, string argumentsJson)
        {
            ShellUniformLoadArgs args = ReadArgs<ShellUniformLoadArgs>(argumentsJson);
            return args.DryRun
                ? Preview(service.PreviewAssignShellUniformLoad(args.AreaNames, args.LoadPattern, args.Value, args.Direction, args.Replace, args.CoordinateSystem))
                : Result(service.AssignShellUniformLoad(args.AreaNames, args.LoadPattern, args.Value, args.Direction, args.Replace, args.CoordinateSystem, args.Confirmed));
        }
    }

    public sealed class ShellsDeleteTool : CsiShellToolBase
    {
        public ShellsDeleteTool(ICSISapModelConnectionService etabsService, ICSISapModelConnectionService sap2000Service) : base(etabsService, sap2000Service) { }
        public override string Name => "shells.delete";
        public override string Title => "Delete Shell / Area Objects";
        public override string SubCategory => "Deletion";
        public override string Description => "Deletes shell/area objects. High risk and requires explicit confirmation.";
        public override bool IsReadOnly => false;
        public override CsiMethodRiskLevel RiskLevel => CsiMethodRiskLevel.High;
        public override bool RequiresConfirmation => true;
        public override bool SupportsDryRun => true;

        protected override ToolCallResponse Execute(ICSISapModelConnectionService service, string argumentsJson)
        {
            ShellDeleteArgs args = ReadArgs<ShellDeleteArgs>(argumentsJson);
            return args.DryRun
                ? Preview(service.PreviewDeleteShells(args.AreaNames))
                : Result(service.DeleteShells(args.AreaNames, args.Confirmed));
        }
    }

    public sealed class ShellNameArgs
    {
        public string AreaName { get; set; }
    }

    public sealed class ShellByPointsArgs : LowRiskWriteArgs
    {
        public List<string> PointNames { get; set; }
        public string PropertyName { get; set; }
        public string UserName { get; set; }
    }

    public sealed class ShellByCoordinatesArgs : LowRiskWriteArgs
    {
        public List<CSISapModelShellCoordinateInput> Points { get; set; }
        public string PropertyName { get; set; }
        public string UserName { get; set; }
        public string CoordinateSystem { get; set; }
    }

    public sealed class ShellUniformLoadArgs : DryRunConfirmedArgs
    {
        public List<string> AreaNames { get; set; }
        public string LoadPattern { get; set; }
        public double Value { get; set; }
        public int Direction { get; set; }
        public bool Replace { get; set; } = true;
        public string CoordinateSystem { get; set; }
    }

    public sealed class ShellDeleteArgs : DryRunConfirmedArgs
    {
        public List<string> AreaNames { get; set; }
    }
}
