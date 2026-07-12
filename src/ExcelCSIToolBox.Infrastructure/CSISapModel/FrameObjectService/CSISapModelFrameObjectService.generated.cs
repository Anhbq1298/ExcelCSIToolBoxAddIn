using System;
using System.Collections.Generic;
using ExcelCSIToolBox.Core.Common.Results;
using ExcelCSIToolBox.Data.CSISapModel.FrameObject;
using ExcelCSIToolBox.Data.DTOs.CSI;
using ExcelCSIToolBox.Data.Models;

namespace ExcelCSIToolBox.Infrastructure.CSISapModel
{
    // AUTO-GENERATED FILE. DO NOT EDIT MANUALLY.
    // Generated scaffold from CSI API reference metadata. Keep safety/business logic in the manual companion file.
    internal static partial class CSISapModelFrameObjectService
    {
        internal delegate int CSISapModelSetFrameSectionGenerated<TSapModel>(TSapModel sapModel, string name, string propertyName);
        internal delegate int CSISapModelDeleteFrameGenerated<TSapModel>(TSapModel sapModel, string name);

        internal static OperationResult<IReadOnlyList<string>> GetNameListGenerated<TSapModel>(
            string productName,
            TSapModel sapModel,
            CSISapModelGetNameList<TSapModel> getNameList)
        {
            return GetNameList(productName, sapModel, getNameList);
        }

        internal static OperationResult<FrameEndPointInfo> GetPointsGenerated<TSapModel>(
            string productName,
            TSapModel sapModel,
            string frameName,
            CSISapModelReadFramePoints<TSapModel> getPoints)
        {
            return GetPoints(productName, sapModel, frameName, getPoints);
        }

        internal static OperationResult<FrameSectionInfo> GetSectionGenerated<TSapModel>(
            string productName,
            TSapModel sapModel,
            string frameName,
            CSISapModelReadFrameSection<TSapModel> getSection)
        {
            return GetSection(productName, sapModel, frameName, getSection);
        }

        internal static OperationResult<CSISapModelAddFramesResultDTO> AddByCoordGenerated<TSapModel>(
            string productName,
            TSapModel sapModel,
            CSISapModelFrameByCoordInput frameInput,
            CSISapModelAddFrameByCoord<TSapModel> addFrame,
            Func<TSapModel, OperationResult> refreshView)
        {
            return AddFramesByCoordinates(new[] { frameInput }, productName, sapModel, addFrame, refreshView);
        }

        internal static OperationResult<CSISapModelAddFramesResultDTO> AddByPointGenerated<TSapModel>(
            string productName,
            TSapModel sapModel,
            CSISapModelFrameByPointInput frameInput,
            CSISapModelAddFrameByPoint<TSapModel> addFrame,
            Func<TSapModel, OperationResult> refreshView)
        {
            return AddFramesByPoint(new[] { frameInput }, productName, sapModel, addFrame, refreshView);
        }

        internal static OperationResult SetSectionGenerated<TSapModel>(
            string productName,
            TSapModel sapModel,
            string frameName,
            string propertyName,
            CSISapModelSetFrameSectionGenerated<TSapModel> setSection)
        {
            if (string.IsNullOrWhiteSpace(frameName) || string.IsNullOrWhiteSpace(propertyName))
            {
                return OperationResult.Failure("Frame name and section property are required.");
            }

            int result = setSection(sapModel, frameName.Trim(), propertyName.Trim());
            return result == 0
                ? OperationResult.Success($"Assigned section '{propertyName}' to {productName} frame '{frameName}'.")
                : OperationResult.Failure($"{productName} FrameObj.SetSection failed for '{frameName}' (return code {result}).");
        }

        internal static OperationResult DeleteGenerated<TSapModel>(
            string productName,
            TSapModel sapModel,
            string frameName,
            CSISapModelDeleteFrameGenerated<TSapModel> deleteFrame)
        {
            if (string.IsNullOrWhiteSpace(frameName))
            {
                return OperationResult.Failure("Frame name is required.");
            }

            int result = deleteFrame(sapModel, frameName.Trim());
            return result == 0
                ? OperationResult.Success($"Deleted {productName} frame '{frameName}'.")
                : OperationResult.Failure($"{productName} FrameObj.Delete failed for '{frameName}' (return code {result}).");
        }
    }
}
