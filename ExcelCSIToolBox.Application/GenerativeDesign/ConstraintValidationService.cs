using ExcelCSIToolBox.Core.Common.Results;
using ExcelCSIToolBox.Core.GenerativeDesign;

namespace ExcelCSIToolBox.Application.GenerativeDesign
{
    public sealed class ConstraintValidationService
    {
        public OperationResult Validate(DesignConstraintSet constraints)
        {
            if (constraints == null)
            {
                return OperationResult.Success("Default design constraints will be used.");
            }

            if (constraints.MinSpan < 0 || constraints.MaxSpan < 0)
            {
                return OperationResult.Failure("Span constraints cannot be negative.");
            }

            if (constraints.MinSpan > 0 && constraints.MaxSpan > 0 && constraints.MinSpan > constraints.MaxSpan)
            {
                return OperationResult.Failure("Minimum span cannot exceed maximum span.");
            }

            if (constraints.MinStories < 0 || constraints.MaxStories < 0)
            {
                return OperationResult.Failure("Story constraints cannot be negative.");
            }

            if (constraints.MinStories > 0 && constraints.MaxStories > 0 && constraints.MinStories > constraints.MaxStories)
            {
                return OperationResult.Failure("Minimum story count cannot exceed maximum story count.");
            }

            return OperationResult.Success("Design constraints are valid.");
        }
    }
}
