using System;
using System.Collections.Generic;
using ExcelCSIToolBox.Core.Abstractions.CSI;
using ExcelCSIToolBox.Core.Common.Results;
using ExcelCSIToolBox.Core.Models.CSI;

namespace ExcelCSIToolBox.Infrastructure.CSISapModel
{
    public sealed class CsiWriteGuard : IMcpWriteGuard
    {
        private static readonly HashSet<string> BlockedByDefault =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "file.save_model",
                "File.Save",
                "File.OpenFile",
                "File.NewBlank",
                "SetModelIsLocked"
            };

        public OperationResult ValidateWrite(
            string operationName,
            CsiMethodRiskLevel riskLevel,
            bool confirmed,
            IReadOnlyList<string> affectedObjects)
        {
            if (string.IsNullOrWhiteSpace(operationName))
            {
                return OperationResult.Failure("Operation name is required.");
            }

            if (IsBlockedByDefault(operationName) || riskLevel == CsiMethodRiskLevel.Dangerous)
            {
                return OperationResult.Failure(
                    $"Operation '{operationName}' is blocked by default because it can permanently alter files or model state.");
            }

            if ((riskLevel == CsiMethodRiskLevel.Medium || riskLevel == CsiMethodRiskLevel.High) && !confirmed)
            {
                return OperationResult.Failure(
                    $"Operation '{operationName}' requires explicit user confirmation before execution.");
            }

            if (riskLevel == CsiMethodRiskLevel.High &&
                (affectedObjects == null || affectedObjects.Count == 0))
            {
                return OperationResult.Failure(
                    $"Operation '{operationName}' is high risk and must name affected objects.");
            }

            return OperationResult.Success("Write guard approved.");
        }

        public bool IsBlockedByDefault(string operationName)
        {
            return !string.IsNullOrWhiteSpace(operationName) &&
                   BlockedByDefault.Contains(operationName);
        }
    }
}
