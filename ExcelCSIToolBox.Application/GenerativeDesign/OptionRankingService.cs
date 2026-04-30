using System.Collections.Generic;
using ExcelCSIToolBox.Core.Common.Results;
using ExcelCSIToolBox.Core.GenerativeDesign;

namespace ExcelCSIToolBox.Application.GenerativeDesign
{
    public sealed class OptionRankingService
    {
        public OperationResult<IReadOnlyList<EvaluationResult>> Rank(IReadOnlyList<EvaluationResult> evaluations)
        {
            var ranked = new List<EvaluationResult>();
            if (evaluations != null)
            {
                ranked.AddRange(evaluations);
            }

            ranked.Sort(Compare);
            return OperationResult<IReadOnlyList<EvaluationResult>>.Success(ranked);
        }

        private static int Compare(EvaluationResult left, EvaluationResult right)
        {
            if (left == null && right == null)
            {
                return 0;
            }

            if (left == null)
            {
                return 1;
            }

            if (right == null)
            {
                return -1;
            }

            int validComparison = right.IsValid.CompareTo(left.IsValid);
            return validComparison != 0
                ? validComparison
                : right.Score.CompareTo(left.Score);
        }
    }
}
