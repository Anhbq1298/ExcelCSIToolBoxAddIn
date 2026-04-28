using System.Collections.Generic;
using ExcelCSIToolBox.Core.Common.Results;
using ExcelCSIToolBox.Core.Models.CSI;

namespace ExcelCSIToolBox.Core.Abstractions.CSI
{
    public interface IMcpWriteGuard
    {
        OperationResult ValidateWrite(
            string operationName,
            CsiMethodRiskLevel riskLevel,
            bool confirmed,
            IReadOnlyList<string> affectedObjects);

        bool IsBlockedByDefault(string operationName);
    }
}
