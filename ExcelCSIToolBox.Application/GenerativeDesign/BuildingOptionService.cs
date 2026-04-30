using System;
using System.Collections.Generic;
using ExcelCSIToolBox.Core.Common.Results;
using ExcelCSIToolBox.Core.GenerativeDesign;

namespace ExcelCSIToolBox.Application.GenerativeDesign
{
    public sealed class BuildingOptionService
    {
        public OperationResult<IReadOnlyList<BuildingOption>> GenerateOptions(DesignConstraintSet constraints, int optionCount)
        {
            DesignConstraintSet normalized = Normalize(constraints);
            int count = Math.Max(1, Math.Min(optionCount <= 0 ? 3 : optionCount, 12));
            var options = new List<BuildingOption>();

            for (int i = 0; i < count; i++)
            {
                int bayCountX = 3 + (i % 4);
                int bayCountY = 2 + (i % 3);
                int stories = Clamp(normalized.MinStories + i, normalized.MinStories, normalized.MaxStories);
                double span = Interpolate(normalized.MinSpan, normalized.MaxSpan, count == 1 ? 0 : (double)i / (count - 1));

                var scheme = new StructuralScheme
                {
                    SchemeType = i % 2 == 0 ? "MomentFrame" : "BracedFrame",
                    BayCountX = bayCountX,
                    BayCountY = bayCountY,
                    StoryCount = stories,
                    TypicalBayLength = span / bayCountX,
                    TypicalStoryHeight = 3600,
                    PrimaryMaterials = new List<string>(normalized.PreferredMaterials)
                };

                if (scheme.PrimaryMaterials.Count == 0)
                {
                    scheme.PrimaryMaterials.Add(i % 2 == 0 ? "Steel" : "Concrete");
                }

                double estimatedWeight = bayCountX * bayCountY * Math.Max(1, stories) * (scheme.SchemeType == "MomentFrame" ? 8.5 : 7.2);
                options.Add(new BuildingOption
                {
                    OptionId = "OPT-" + (i + 1).ToString("000"),
                    Name = scheme.SchemeType + " " + bayCountX + "x" + bayCountY + "x" + stories,
                    Scheme = scheme,
                    EstimatedWeight = estimatedWeight,
                    EstimatedCost = estimatedWeight * 1250,
                    Notes = "Conceptual option generated for MCP preview and ranking."
                });
            }

            return OperationResult<IReadOnlyList<BuildingOption>>.Success(options);
        }

        public OperationResult<BuildingOption> PreviewOption(BuildingOption option)
        {
            if (option == null)
            {
                return OperationResult<BuildingOption>.Failure("A building option is required.");
            }

            return OperationResult<BuildingOption>.Success(option);
        }

        public OperationResult<BuildingOption> BuildOption(BuildingOption option, bool dryRun, bool confirmed)
        {
            if (option == null)
            {
                return OperationResult<BuildingOption>.Failure("A building option is required.");
            }

            if (dryRun)
            {
                return OperationResult<BuildingOption>.Success(option, "Preview only. No CSI model changes were made.");
            }

            if (!confirmed)
            {
                return OperationResult<BuildingOption>.Failure("Building an option requires explicit confirmation.");
            }

            return OperationResult<BuildingOption>.Success(
                option,
                "Build request accepted. Low-level CSI generation is intentionally delegated to application workflows.");
        }

        private static DesignConstraintSet Normalize(DesignConstraintSet constraints)
        {
            DesignConstraintSet source = constraints ?? new DesignConstraintSet();
            double minSpan = source.MinSpan > 0 ? source.MinSpan : 18000;
            double maxSpan = source.MaxSpan > 0 ? source.MaxSpan : 36000;
            if (maxSpan < minSpan)
            {
                double temp = minSpan;
                minSpan = maxSpan;
                maxSpan = temp;
            }

            int minStories = source.MinStories > 0 ? source.MinStories : 1;
            int maxStories = source.MaxStories > 0 ? source.MaxStories : Math.Max(minStories, 5);

            return new DesignConstraintSet
            {
                MinSpan = minSpan,
                MaxSpan = maxSpan,
                MinStories = Math.Min(minStories, maxStories),
                MaxStories = Math.Max(minStories, maxStories),
                MaxDriftRatio = source.MaxDriftRatio > 0 ? source.MaxDriftRatio : 1.0 / 400.0,
                MaxWeight = source.MaxWeight,
                PreferredMaterials = source.PreferredMaterials == null
                    ? new List<string>()
                    : new List<string>(source.PreferredMaterials)
            };
        }

        private static double Interpolate(double min, double max, double ratio)
        {
            return min + (max - min) * ratio;
        }

        private static int Clamp(int value, int min, int max)
        {
            return Math.Max(min, Math.Min(max, value));
        }
    }
}
