using ExcelCSIToolBox.Core.Common.Results;
using ExcelCSIToolBox.Core.GenerativeDesign;

namespace ExcelCSIToolBox.Application.GenerativeDesign
{
    public sealed class ResultEvaluationService
    {
        public OperationResult<EvaluationResult> Evaluate(BuildingOption option, DesignConstraintSet constraints)
        {
            if (option == null)
            {
                return OperationResult<EvaluationResult>.Failure("A building option is required.");
            }

            DesignConstraintSet normalized = constraints ?? new DesignConstraintSet();
            double driftLimit = normalized.MaxDriftRatio > 0 ? normalized.MaxDriftRatio : 1.0 / 400.0;
            double estimatedDrift = EstimateDrift(option);
            bool weightOk = normalized.MaxWeight <= 0 || option.EstimatedWeight <= normalized.MaxWeight;
            bool driftOk = estimatedDrift <= driftLimit;

            var result = new EvaluationResult
            {
                OptionId = option.OptionId,
                IsValid = weightOk && driftOk,
                EstimatedDriftRatio = estimatedDrift,
                EstimatedWeight = option.EstimatedWeight,
                Score = CalculateScore(option, estimatedDrift, driftLimit, weightOk)
            };

            result.Messages.Add(driftOk ? "Estimated drift is within limit." : "Estimated drift exceeds limit.");
            result.Messages.Add(weightOk ? "Estimated weight is within limit." : "Estimated weight exceeds limit.");

            return OperationResult<EvaluationResult>.Success(result);
        }

        private static double EstimateDrift(BuildingOption option)
        {
            if (option.Scheme == null || option.Scheme.StoryCount <= 0)
            {
                return 1.0 / 250.0;
            }

            double systemFactor = option.Scheme.SchemeType == "BracedFrame" ? 0.75 : 1.0;
            return systemFactor * option.Scheme.StoryCount / 2500.0;
        }

        private static double CalculateScore(BuildingOption option, double estimatedDrift, double driftLimit, bool weightOk)
        {
            double driftScore = driftLimit <= 0 ? 50 : 60.0 * System.Math.Min(1.0, driftLimit / System.Math.Max(estimatedDrift, 0.000001));
            double weightScore = weightOk ? 30.0 : 10.0;
            double simplicityScore = option.Scheme == null ? 0 : System.Math.Max(0, 10 - option.Scheme.BayCountX);
            return driftScore + weightScore + simplicityScore;
        }
    }
}
