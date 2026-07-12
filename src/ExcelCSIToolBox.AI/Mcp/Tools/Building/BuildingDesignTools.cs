using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using ExcelCSIToolBox.AI.Mcp.Contracts;
using ExcelCSIToolBox.Application.GenerativeDesign;
using ExcelCSIToolBox.Core.Common.Results;
using ExcelCSIToolBox.Core.GenerativeDesign;
using ExcelCSIToolBox.Core.Models.CSI;
using Newtonsoft.Json;

namespace ExcelCSIToolBox.AI.Mcp.Tools.Building
{
    public abstract class BuildingDesignToolBase : IMcpTool, IMcpToolMetadata
    {
        protected BuildingDesignToolBase(
            BuildingOptionService buildingOptionService,
            ConstraintValidationService constraintValidationService,
            ResultEvaluationService resultEvaluationService,
            OptionRankingService optionRankingService)
        {
            BuildingOptionService = buildingOptionService;
            ConstraintValidationService = constraintValidationService;
            ResultEvaluationService = resultEvaluationService;
            OptionRankingService = optionRankingService;
        }

        protected BuildingOptionService BuildingOptionService { get; }
        protected ConstraintValidationService ConstraintValidationService { get; }
        protected ResultEvaluationService ResultEvaluationService { get; }
        protected OptionRankingService OptionRankingService { get; }

        public abstract string Name { get; }
        public abstract string Title { get; }
        public string Category => "Building";
        public abstract string SubCategory { get; }
        public abstract string Description { get; }
        public virtual bool IsReadOnly => true;
        public virtual CsiMethodRiskLevel RiskLevel => CsiMethodRiskLevel.None;
        public virtual bool RequiresConfirmation => false;
        public virtual bool SupportsDryRun => false;

        public Task<ToolCallResponse> ExecuteAsync(string argumentsJson, CancellationToken cancellationToken)
        {
            try
            {
                return Task.FromResult(Execute(argumentsJson ?? "{}"));
            }
            catch (System.Exception ex)
            {
                return Task.FromResult(Fail("Building design tool failed: " + ex.Message));
            }
        }

        protected abstract ToolCallResponse Execute(string argumentsJson);

        protected static TArgs ReadArgs<TArgs>(string argumentsJson) where TArgs : class, new()
        {
            return JsonConvert.DeserializeObject<TArgs>(argumentsJson ?? "{}") ?? new TArgs();
        }

        protected ToolCallResponse Result<T>(OperationResult<T> result)
        {
            return result.IsSuccess
                ? Ok(result.Message ?? "Building design task completed.", result.Data)
                : Fail(result.Message);
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
    }

    public sealed class BuildingGenerateOptionsTool : BuildingDesignToolBase
    {
        public BuildingGenerateOptionsTool(BuildingOptionService buildingOptionService, ConstraintValidationService constraintValidationService, ResultEvaluationService resultEvaluationService, OptionRankingService optionRankingService)
            : base(buildingOptionService, constraintValidationService, resultEvaluationService, optionRankingService) { }

        public override string Name => "building.generate_options";
        public override string Title => "Generate Building Options";
        public override string SubCategory => "Generation";
        public override string Description => "Generates conceptual building options from constraints without mutating a CSI model.";

        protected override ToolCallResponse Execute(string argumentsJson)
        {
            GenerateOptionsArgs args = ReadArgs<GenerateOptionsArgs>(argumentsJson);
            OperationResult validation = ConstraintValidationService.Validate(args.Constraints);
            if (!validation.IsSuccess)
            {
                return Fail(validation.Message);
            }

            return Result(BuildingOptionService.GenerateOptions(args.Constraints, args.OptionCount));
        }
    }

    public sealed class BuildingPreviewOptionTool : BuildingDesignToolBase
    {
        public BuildingPreviewOptionTool(BuildingOptionService buildingOptionService, ConstraintValidationService constraintValidationService, ResultEvaluationService resultEvaluationService, OptionRankingService optionRankingService)
            : base(buildingOptionService, constraintValidationService, resultEvaluationService, optionRankingService) { }

        public override string Name => "building.preview_option";
        public override string Title => "Preview Building Option";
        public override string SubCategory => "Preview";
        public override string Description => "Previews a conceptual building option without writing to a CSI model.";

        protected override ToolCallResponse Execute(string argumentsJson)
        {
            BuildingOptionArgs args = ReadArgs<BuildingOptionArgs>(argumentsJson);
            return Result(BuildingOptionService.PreviewOption(args.Option));
        }
    }

    public sealed class BuildingBuildOptionTool : BuildingDesignToolBase
    {
        public BuildingBuildOptionTool(BuildingOptionService buildingOptionService, ConstraintValidationService constraintValidationService, ResultEvaluationService resultEvaluationService, OptionRankingService optionRankingService)
            : base(buildingOptionService, constraintValidationService, resultEvaluationService, optionRankingService) { }

        public override string Name => "building.build_option";
        public override string Title => "Build Building Option";
        public override string SubCategory => "Build";
        public override string Description => "Accepts a confirmed building option build request through the application layer.";
        public override bool IsReadOnly => false;
        public override CsiMethodRiskLevel RiskLevel => CsiMethodRiskLevel.High;
        public override bool RequiresConfirmation => true;
        public override bool SupportsDryRun => true;

        protected override ToolCallResponse Execute(string argumentsJson)
        {
            BuildingOptionWriteArgs args = ReadArgs<BuildingOptionWriteArgs>(argumentsJson);
            if (args.DryRun)
            {
                return Ok("Preview: building option is ready to build. Confirm to proceed?", new
                {
                    OperationName = Name,
                    Summary = "Build conceptual option " + (args.Option == null ? "(unnamed)" : args.Option.Name) + ".",
                    RequiresConfirmation = true,
                    SupportsDryRun = true,
                    Option = args.Option
                });
            }

            return Result(BuildingOptionService.BuildOption(args.Option, args.DryRun, args.Confirmed));
        }
    }

    public sealed class BuildingRunAnalysisTool : BuildingDesignToolBase
    {
        public BuildingRunAnalysisTool(BuildingOptionService buildingOptionService, ConstraintValidationService constraintValidationService, ResultEvaluationService resultEvaluationService, OptionRankingService optionRankingService)
            : base(buildingOptionService, constraintValidationService, resultEvaluationService, optionRankingService) { }

        public override string Name => "building.run_analysis";
        public override string Title => "Run Conceptual Building Analysis";
        public override string SubCategory => "Analysis";
        public override string Description => "Runs conceptual application-level analysis for a building option.";
        public override CsiMethodRiskLevel RiskLevel => CsiMethodRiskLevel.Medium;
        public override bool SupportsDryRun => true;

        protected override ToolCallResponse Execute(string argumentsJson)
        {
            EvaluateOptionArgs args = ReadArgs<EvaluateOptionArgs>(argumentsJson);
            return Result(ResultEvaluationService.Evaluate(args.Option, args.Constraints));
        }
    }

    public sealed class BuildingEvaluateOptionTool : BuildingDesignToolBase
    {
        public BuildingEvaluateOptionTool(BuildingOptionService buildingOptionService, ConstraintValidationService constraintValidationService, ResultEvaluationService resultEvaluationService, OptionRankingService optionRankingService)
            : base(buildingOptionService, constraintValidationService, resultEvaluationService, optionRankingService) { }

        public override string Name => "building.evaluate_option";
        public override string Title => "Evaluate Building Option";
        public override string SubCategory => "Evaluation";
        public override string Description => "Evaluates a conceptual building option against constraints.";

        protected override ToolCallResponse Execute(string argumentsJson)
        {
            EvaluateOptionArgs args = ReadArgs<EvaluateOptionArgs>(argumentsJson);
            return Result(ResultEvaluationService.Evaluate(args.Option, args.Constraints));
        }
    }

    public sealed class BuildingRankOptionsTool : BuildingDesignToolBase
    {
        public BuildingRankOptionsTool(BuildingOptionService buildingOptionService, ConstraintValidationService constraintValidationService, ResultEvaluationService resultEvaluationService, OptionRankingService optionRankingService)
            : base(buildingOptionService, constraintValidationService, resultEvaluationService, optionRankingService) { }

        public override string Name => "building.rank_options";
        public override string Title => "Rank Building Options";
        public override string SubCategory => "Ranking";
        public override string Description => "Ranks evaluated building options.";

        protected override ToolCallResponse Execute(string argumentsJson)
        {
            RankOptionsArgs args = ReadArgs<RankOptionsArgs>(argumentsJson);
            return Result(OptionRankingService.Rank(args.Evaluations));
        }
    }

    internal sealed class GenerateOptionsArgs
    {
        public DesignConstraintSet Constraints { get; set; }
        public int OptionCount { get; set; }
    }

    internal class BuildingOptionArgs
    {
        public BuildingOption Option { get; set; }
    }

    internal sealed class BuildingOptionWriteArgs : BuildingOptionArgs
    {
        public bool DryRun { get; set; } = true;
        public bool Confirmed { get; set; }
    }

    internal sealed class EvaluateOptionArgs : BuildingOptionArgs
    {
        public DesignConstraintSet Constraints { get; set; }
    }

    internal sealed class RankOptionsArgs
    {
        public List<EvaluationResult> Evaluations { get; set; }
    }
}
