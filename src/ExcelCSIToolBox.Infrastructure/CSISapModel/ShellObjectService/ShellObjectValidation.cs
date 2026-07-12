using System.Collections.Generic;
using ExcelCSIToolBox.Core.Common.Results;
using ExcelCSIToolBox.Data.Models;

namespace ExcelCSIToolBox.Infrastructure.CSISapModel
{
    internal static class ShellObjectValidation
    {
        internal static OperationResult ValidateAreaName(string areaName)
        {
            return string.IsNullOrWhiteSpace(areaName)
                ? OperationResult.Failure("Area name is required.")
                : OperationResult.Success();
        }

        internal static OperationResult ValidateAreaNames(IReadOnlyList<string> areaNames)
        {
            if (areaNames == null || areaNames.Count == 0)
            {
                return OperationResult.Failure("At least one area name is required.");
            }

            for (int i = 0; i < areaNames.Count; i++)
            {
                if (string.IsNullOrWhiteSpace(areaNames[i]))
                {
                    return OperationResult.Failure("Area names cannot contain empty values.");
                }
            }

            return OperationResult.Success();
        }

        internal static OperationResult ValidatePointNames(IReadOnlyList<string> pointNames)
        {
            if (pointNames == null || pointNames.Count < 3)
            {
                return OperationResult.Failure("At least three point names are required to create a shell/area object.");
            }

            for (int i = 0; i < pointNames.Count; i++)
            {
                if (string.IsNullOrWhiteSpace(pointNames[i]))
                {
                    return OperationResult.Failure("Point names cannot contain empty values.");
                }
            }

            return OperationResult.Success();
        }

        internal static OperationResult ValidateCoordinates(IReadOnlyList<CSISapModelShellCoordinateInput> points)
        {
            if (points == null || points.Count < 3)
            {
                return OperationResult.Failure("At least three coordinates are required to create a shell/area object.");
            }

            for (int i = 0; i < points.Count; i++)
            {
                if (points[i] == null)
                {
                    return OperationResult.Failure("Coordinate list cannot contain null points.");
                }
            }

            return OperationResult.Success();
        }

        internal static OperationResult ValidateUniformLoad(string loadPattern, int direction, string coordinateSystem)
        {
            if (string.IsNullOrWhiteSpace(loadPattern))
            {
                return OperationResult.Failure("Load pattern is required.");
            }

            if (direction < 1 || direction > 11)
            {
                return OperationResult.Failure("Direction code must be between 1 and 11 for CSI area uniform loads.");
            }

            if (string.IsNullOrWhiteSpace(coordinateSystem))
            {
                return OperationResult.Failure("Coordinate system is required.");
            }

            return OperationResult.Success();
        }
    }
}
