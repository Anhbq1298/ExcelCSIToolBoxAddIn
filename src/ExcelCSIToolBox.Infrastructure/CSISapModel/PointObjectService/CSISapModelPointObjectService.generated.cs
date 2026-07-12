using System;
using System.Collections.Generic;
using ExcelCSIToolBox.Core.Common.Results;
using ExcelCSIToolBox.Data.CSISapModel.PointObject;
using ExcelCSIToolBox.Data.DTOs.CSI;
using ExcelCSIToolBox.Data.Models;

namespace ExcelCSIToolBox.Infrastructure.CSISapModel
{
    // AUTO-GENERATED FILE. DO NOT EDIT MANUALLY.
    // Generated scaffold from CSI API reference metadata. Keep safety/business logic in the manual companion file.
    internal static partial class CSISapModelPointObjectService
    {
        internal delegate int CSISapModelDeletePointGenerated<TSapModel>(TSapModel sapModel, string name);

        internal static OperationResult<IReadOnlyList<string>> GetNameListGenerated<TSapModel>(
            string productName,
            TSapModel sapModel,
            CSISapModelGetNameList<TSapModel> getNameList)
        {
            return GetNameList(productName, sapModel, getNameList);
        }

        internal static OperationResult<PointObjectInfo> GetCoordCartesianGenerated<TSapModel>(
            string productName,
            TSapModel sapModel,
            string pointName,
            CSISapModelReadPointCoordinates<TSapModel> getPointCoordinates)
        {
            return GetByName(productName, sapModel, pointName, getPointCoordinates, null);
        }

        internal static OperationResult<CSISapModelAddPointsResultDTO> AddCartesianGenerated<TSapModel>(
            string productName,
            TSapModel sapModel,
            CSISapModelPointCartesianInput pointInput,
            CSISapModelAddCartesianPoint<TSapModel> addPoint,
            Func<TSapModel, OperationResult> refreshView)
        {
            return AddPointsByCartesian(new[] { pointInput }, productName, sapModel, addPoint, refreshView);
        }

        internal static OperationResult DeleteGenerated<TSapModel>(
            string productName,
            TSapModel sapModel,
            string pointName,
            CSISapModelDeletePointGenerated<TSapModel> deletePoint)
        {
            if (string.IsNullOrWhiteSpace(pointName))
            {
                return OperationResult.Failure("Point name is required.");
            }

            int result = deletePoint(sapModel, pointName.Trim());
            return result == 0
                ? OperationResult.Success($"Deleted {productName} point '{pointName}'.")
                : OperationResult.Failure($"{productName} PointObj.Delete failed for '{pointName}' (return code {result}).");
        }
    }
}
