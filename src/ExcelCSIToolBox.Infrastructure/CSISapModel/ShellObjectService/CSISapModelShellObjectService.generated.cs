using System.Collections.Generic;
using ExcelCSIToolBox.Core.Common.Results;

namespace ExcelCSIToolBox.Infrastructure.CSISapModel
{
    // AUTO-GENERATED FILE. DO NOT EDIT MANUALLY.
    // Generated scaffold from CSI API reference metadata. Keep safety/business logic in the manual companion file.
    internal static partial class CSISapModelShellObjectService
    {
        internal delegate int CSISapModelSetShellPropertyGenerated<TSapModel>(TSapModel sapModel, string name, string propertyName);
        internal delegate int CSISapModelDeleteShellGenerated<TSapModel>(TSapModel sapModel, string name);

        internal static OperationResult<IReadOnlyList<string>> GetNameListGenerated<TSapModel>(
            string productName,
            TSapModel sapModel,
            CSISapModelAreaGetNameList<TSapModel> getNameList)
        {
            return GetNameList<TSapModel>(sapModel, productName, getNameList);
        }

        internal static OperationResult<IReadOnlyList<string>> GetPointsGenerated<TSapModel>(
            string productName,
            TSapModel sapModel,
            string areaName,
            CSISapModelAreaGetPoints<TSapModel> getPoints)
        {
            return GetPoints<TSapModel>(sapModel, productName, areaName, getPoints);
        }

        internal static OperationResult SetPropertyGenerated<TSapModel>(
            string productName,
            TSapModel sapModel,
            string areaName,
            string propertyName,
            CSISapModelSetShellPropertyGenerated<TSapModel> setProperty)
        {
            if (string.IsNullOrWhiteSpace(areaName) || string.IsNullOrWhiteSpace(propertyName))
            {
                return OperationResult.Failure("Shell/area name and property name are required.");
            }

            int result = setProperty(sapModel, areaName.Trim(), propertyName.Trim());
            return result == 0
                ? OperationResult.Success($"Assigned property '{propertyName}' to {productName} shell/area '{areaName}'.")
                : OperationResult.Failure($"{productName} AreaObj.SetProperty failed for '{areaName}' (return code {result}).");
        }

        internal static OperationResult DeleteGenerated<TSapModel>(
            string productName,
            TSapModel sapModel,
            string areaName,
            CSISapModelDeleteShellGenerated<TSapModel> deleteShell)
        {
            if (string.IsNullOrWhiteSpace(areaName))
            {
                return OperationResult.Failure("Shell/area name is required.");
            }

            int result = deleteShell(sapModel, areaName.Trim());
            return result == 0
                ? OperationResult.Success($"Deleted {productName} shell/area '{areaName}'.")
                : OperationResult.Failure($"{productName} AreaObj.Delete failed for '{areaName}' (return code {result}).");
        }
    }
}
