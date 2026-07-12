using System.Collections.Generic;
using ExcelCSIToolBox.Core.Models.CSI;

namespace ExcelCSIToolBox.Core.Abstractions.CSI
{
    public interface ICsiOperationLogger
    {
        void Log(
            string productType,
            string operationName,
            string category,
            string subCategory,
            CsiMethodRiskLevel riskLevel,
            string argumentsSummary,
            IReadOnlyList<string> affectedObjects,
            bool confirmed,
            bool succeeded,
            string message);
    }
}
