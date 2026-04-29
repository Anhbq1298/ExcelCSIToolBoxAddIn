using System.Runtime.InteropServices;
using System;
using System.Collections.Generic;
using ExcelCSIToolBox.Core.Common.Results;
using ExcelCSIToolBox.Data;
using ExcelCSIToolBox.Data.CSISapModel.FrameObject;
using ExcelCSIToolBox.Data.DTOs.CSI;
using ExcelCSIToolBox.Data.Models;

namespace ExcelCSIToolBox.Infrastructure.CSISapModel
{
    public delegate int CSISapModelAddFrameByCoord<TSapModel>(
        TSapModel sapModel,
        CSISapModelFrameByCoordInput frameInput,
        ref string createdName,
        string sectionName,
        string userName);

    public delegate int CSISapModelAddFrameByPoint<TSapModel>(
        TSapModel sapModel,
        CSISapModelFrameByPointInput frameInput,
        ref string createdName,
        string sectionName,
        string userName);

    internal delegate int CSISapModelReadFramePoints<TSapModel>(
        TSapModel sapModel,
        string frameName,
        ref string pointI,
        ref string pointJ);

    internal delegate int CSISapModelReadFrameSection<TSapModel>(
        TSapModel sapModel,
        string frameName,
        ref string sectionName,
        ref string autoSelectList);

    internal delegate int CSISapModelReadFrameSelected<TSapModel>(
        TSapModel sapModel,
        string frameName,
        ref bool selected);

    internal delegate int CSISapModelReadFrameDistributedLoads<TSapModel>(
        TSapModel sapModel,
        string frameName,
        ref int numberItems,
        ref string[] frameNames,
        ref string[] loadPatterns,
        ref int[] loadTypes,
        ref string[] coordinateSystems,
        ref int[] directions,
        ref double[] relativeDistance1,
        ref double[] relativeDistance2,
        ref double[] distance1,
        ref double[] distance2,
        ref double[] value1,
        ref double[] value2);

    internal delegate int CSISapModelReadFramePointLoads<TSapModel>(
        TSapModel sapModel,
        string frameName,
        ref int numberItems,
        ref string[] frameNames,
        ref string[] loadPatterns,
        ref int[] loadTypes,
        ref string[] coordinateSystems,
        ref int[] directions,
        ref double[] relativeDistance,
        ref double[] distance,
        ref double[] value);

    internal static partial class CSISapModelFrameObjectService
    {
        internal static OperationResult SelectFramesByUniqueNames<TSapModel>(
            IReadOnlyList<string> uniqueNames,
            string productName,
            TSapModel sapModel,
            CSISapModelClearSelection<TSapModel> clearSelection,
            CSISapModelSetSelectedByName<TSapModel> setSelected,
            Func<TSapModel, OperationResult> refreshView)
        {
            return SelectObjectsByUniqueNames(
                uniqueNames,
                "frame",
                productName,
                sapModel,
                clearSelection,
                setSelected,
                refreshView);
        }

        internal static OperationResult<CSISapModelAddFramesResultDTO> AddFramesByCoordinates<TSapModel>(
            IReadOnlyList<CSISapModelFrameByCoordInput> frameInputs,
            string productName,
            TSapModel sapModel,
            CSISapModelAddFrameByCoord<TSapModel> addFrame,
            Func<TSapModel, OperationResult> refreshView)
        {
            return AddFrames(
                frameInputs,
                "Adding Frames (by Coordinates)...",
                productName,
                "FrameObj.AddByCoord",
                "coordinates",
                sapModel,
                refreshView,
                (model, frameInput) =>
                {
                    string createdName = string.Empty;
                    string sectionName = string.IsNullOrWhiteSpace(frameInput.SectionName) ? "Default" : frameInput.SectionName;
                    string userName = string.IsNullOrWhiteSpace(frameInput.UniqueName) ? string.Empty : frameInput.UniqueName;
                    return addFrame(model, frameInput, ref createdName, sectionName, userName);
                },
                null);
        }

        internal static OperationResult<CSISapModelAddFramesResultDTO> AddFramesByPoint<TSapModel>(
            IReadOnlyList<CSISapModelFrameByPointInput> frameInputs,
            string productName,
            TSapModel sapModel,
            CSISapModelAddFrameByPoint<TSapModel> addFrame,
            Func<TSapModel, OperationResult> refreshView)
        {
            return AddFrames(
                frameInputs,
                "Adding Frames (by Point Names)...",
                productName,
                "FrameObj.AddByPoint",
                "point names",
                sapModel,
                refreshView,
                (model, frameInput) =>
                {
                    string createdName = string.Empty;
                    string sectionName = string.IsNullOrWhiteSpace(frameInput.SectionName) ? "Default" : frameInput.SectionName;
                    string userName = string.IsNullOrWhiteSpace(frameInput.UniqueName) ? string.Empty : frameInput.UniqueName;
                    int result = addFrame(model, frameInput, ref createdName, sectionName, userName);
                    return result == 0 &&
                           !string.IsNullOrWhiteSpace(userName) &&
                           !string.Equals(createdName, userName, StringComparison.OrdinalIgnoreCase)
                        ? new FrameAddResult(result, $"Frame was created, but {productName} assigned UniqueName '{createdName}' instead of requested '{userName}'.")
                        : new FrameAddResult(result, null);
                },
                result => result.WarningMessage);
        }

        internal static OperationResult<FrameAddBatchResultDto> AddFrameObjects<TSapModel>(
            FrameAddBatchRequestDto request,
            string productName,
            TSapModel sapModel,
            CSISapModelAddFrameByPoint<TSapModel> addByPoint,
            CSISapModelAddFrameByCoord<TSapModel> addByCoord,
            Func<TSapModel, OperationResult> refreshView)
        {
            if (sapModel == null)
            {
                return OperationResult<FrameAddBatchResultDto>.Failure("active CSI model is not available.");
            }

            var frames = request == null ? null : request.Frames;
            if (frames == null || frames.Count == 0)
            {
                return OperationResult<FrameAddBatchResultDto>.Success(new FrameAddBatchResultDto
                {
                    TotalRequested = 0,
                    SuccessCount = 0,
                    FailureCount = 0,
                    SuccessfulFrameNames = new List<string>(),
                    FailedItems = new List<FrameAddResultDto>(),
                    Results = new List<FrameAddResultDto>()
                });
            }

            var batchResult = new FrameAddBatchResultDto
            {
                TotalRequested = frames.Count,
                SuccessfulFrameNames = new List<string>(),
                FailedItems = new List<FrameAddResultDto>(),
                Results = new List<FrameAddResultDto>()
            };

            try
            {
                foreach (FrameAddRequestDto frame in frames)
                {
                    FrameAddResultDto itemResult = AddOneFrameObject(
                        frame,
                        productName,
                        sapModel,
                        addByPoint,
                        addByCoord);

                    batchResult.Results.Add(itemResult);
                    if (itemResult.Success)
                    {
                        batchResult.SuccessCount++;
                        if (!string.IsNullOrWhiteSpace(itemResult.FrameName))
                        {
                            batchResult.SuccessfulFrameNames.Add(itemResult.FrameName);
                        }
                    }
                    else
                    {
                        batchResult.FailureCount++;
                        batchResult.FailedItems.Add(itemResult);
                    }
                }

                if (batchResult.SuccessCount > 0)
                {
                    OperationResult refreshResult = refreshView(sapModel);
                    if (!refreshResult.IsSuccess)
                    {
                        return OperationResult<FrameAddBatchResultDto>.Failure(refreshResult.Message);
                    }
                }

                return OperationResult<FrameAddBatchResultDto>.Success(batchResult);
            }
            catch (COMException ex)
            {
                return OperationResult<FrameAddBatchResultDto>.Failure($"{productName} COM error while adding frame objects: {ex.Message}");
            }
            catch (Exception ex)
            {
                return OperationResult<FrameAddBatchResultDto>.Failure($"{productName} frame add failed unexpectedly: {ex.Message}");
            }
        }

        internal static OperationResult<IReadOnlyList<string>> GetSelectedFramesFromActiveModel<TSapModel>(
            string productName,
            TSapModel sapModel,
            CSISapModelGetSelectedObjects<TSapModel> getSelectedObjects)
        {
            try
            {
                int numberItems = 0;
                int[] objectTypes = null;
                string[] objectNames = null;
                int selectedResult = getSelectedObjects(sapModel, ref numberItems, ref objectTypes, ref objectNames);
                if (selectedResult != 0)
                {
                    return OperationResult<IReadOnlyList<string>>.Failure($"Failed to read selected objects from {productName}.");
                }

                var frameUniqueNames = new List<string>();
                for (int i = 0; i < numberItems; i++)
                {
                    if (objectTypes == null || objectNames == null || i >= objectTypes.Length || i >= objectNames.Length)
                    {
                        continue;
                    }

                    var frameUniqueName = objectNames[i];
                    if (objectTypes[i] == CSISapModelObjectTypeIds.Frame && !string.IsNullOrWhiteSpace(frameUniqueName))
                    {
                        frameUniqueNames.Add(frameUniqueName);
                    }
                }

                if (frameUniqueNames.Count == 0)
                {
                    return OperationResult<IReadOnlyList<string>>.Failure($"No frame objects are selected in {productName}.");
                }

                return OperationResult<IReadOnlyList<string>>.Success(frameUniqueNames);
            }
            catch
            {
                return OperationResult<IReadOnlyList<string>>.Failure($"Unable to read selected {productName} frames.");
            }
        }

        internal static OperationResult<IReadOnlyList<string>> GetNameList<TSapModel>(
            string productName,
            TSapModel sapModel,
            CSISapModelGetNameList<TSapModel> getNameList)
        {
            int numberNames = 0;
            string[] names = null;
            int result = getNameList(sapModel, ref numberNames, ref names);
            if (result != 0)
            {
                return OperationResult<IReadOnlyList<string>>.Failure($"{productName} FrameObj.GetNameList failed (return code {result}).");
            }

            return OperationResult<IReadOnlyList<string>>.Success(names ?? new string[0]);
        }

        internal static OperationResult<int> GetCount<TSapModel>(
            string productName,
            TSapModel sapModel,
            CSISapModelReadCount<TSapModel> getCount)
        {
            int count = 0;
            int result = getCount(sapModel, ref count);
            if (result != 0)
            {
                return OperationResult<int>.Failure($"{productName} FrameObj.Count failed (return code {result}).");
            }

            return OperationResult<int>.Success(count);
        }

        internal static OperationResult<FrameEndPointInfo> GetPoints<TSapModel>(
            string productName,
            TSapModel sapModel,
            string frameName,
            CSISapModelReadFramePoints<TSapModel> getPoints)
        {
            if (string.IsNullOrWhiteSpace(frameName))
            {
                return OperationResult<FrameEndPointInfo>.Failure("Frame name is required.");
            }

            string pointI = string.Empty;
            string pointJ = string.Empty;
            int result = getPoints(sapModel, frameName, ref pointI, ref pointJ);
            if (result != 0)
            {
                return OperationResult<FrameEndPointInfo>.Failure($"{productName} FrameObj.GetPoints failed for '{frameName}' (return code {result}).");
            }

            return OperationResult<FrameEndPointInfo>.Success(new FrameEndPointInfo
            {
                FrameName = frameName,
                PointI = pointI,
                PointJ = pointJ
            });
        }

        internal static OperationResult<FrameSectionInfo> GetSection<TSapModel>(
            string productName,
            TSapModel sapModel,
            string frameName,
            CSISapModelReadFrameSection<TSapModel> getSection)
        {
            if (string.IsNullOrWhiteSpace(frameName))
            {
                return OperationResult<FrameSectionInfo>.Failure("Frame name is required.");
            }

            string sectionName = string.Empty;
            string autoSelectList = string.Empty;
            int result = getSection(sapModel, frameName, ref sectionName, ref autoSelectList);
            if (result != 0)
            {
                return OperationResult<FrameSectionInfo>.Failure($"{productName} FrameObj.GetSection failed for '{frameName}' (return code {result}).");
            }

            return OperationResult<FrameSectionInfo>.Success(new FrameSectionInfo
            {
                FrameName = frameName,
                SectionName = sectionName,
                AutoSelectList = autoSelectList
            });
        }

        internal static OperationResult<FrameObjectInfo> GetByName<TSapModel>(
            string productName,
            TSapModel sapModel,
            string frameName,
            CSISapModelReadFramePoints<TSapModel> getPoints,
            CSISapModelReadFrameSection<TSapModel> getSection,
            CSISapModelReadFrameSelected<TSapModel> getSelected)
        {
            var pointsResult = GetPoints(productName, sapModel, frameName, getPoints);
            if (!pointsResult.IsSuccess)
            {
                return OperationResult<FrameObjectInfo>.Failure(pointsResult.Message);
            }

            var sectionResult = GetSection(productName, sapModel, frameName, getSection);
            if (!sectionResult.IsSuccess)
            {
                return OperationResult<FrameObjectInfo>.Failure(sectionResult.Message);
            }

            bool selected = false;
            if (getSelected != null)
            {
                getSelected(sapModel, frameName, ref selected);
            }

            return OperationResult<FrameObjectInfo>.Success(new FrameObjectInfo
            {
                Name = frameName,
                PointI = pointsResult.Data.PointI,
                PointJ = pointsResult.Data.PointJ,
                SectionName = sectionResult.Data.SectionName,
                IsSelected = selected
            });
        }

        internal static OperationResult<IReadOnlyList<FrameLoadInfo>> GetDistributedLoads<TSapModel>(
            string productName,
            TSapModel sapModel,
            string frameName,
            CSISapModelReadFrameDistributedLoads<TSapModel> getLoadDistributed)
        {
            if (string.IsNullOrWhiteSpace(frameName))
            {
                return OperationResult<IReadOnlyList<FrameLoadInfo>>.Failure("Frame name is required.");
            }

            int numberItems = 0;
            string[] frameNames = null;
            string[] loadPatterns = null;
            int[] loadTypes = null;
            string[] coordinateSystems = null;
            int[] directions = null;
            double[] rd1 = null;
            double[] rd2 = null;
            double[] dist1 = null;
            double[] dist2 = null;
            double[] val1 = null;
            double[] val2 = null;

            int result = getLoadDistributed(sapModel, frameName, ref numberItems, ref frameNames, ref loadPatterns, ref loadTypes, ref coordinateSystems, ref directions, ref rd1, ref rd2, ref dist1, ref dist2, ref val1, ref val2);
            if (result != 0)
            {
                return OperationResult<IReadOnlyList<FrameLoadInfo>>.Failure($"{productName} FrameObj.GetLoadDistributed failed for '{frameName}' (return code {result}).");
            }

            var loads = new List<FrameLoadInfo>();
            for (int i = 0; i < numberItems; i++)
            {
                loads.Add(new FrameLoadInfo
                {
                    FrameName = GetArrayValue(frameNames, i, frameName),
                    LoadPattern = GetArrayValue(loadPatterns, i, string.Empty),
                    LoadType = "Distributed",
                    CoordinateSystem = GetArrayValue(coordinateSystems, i, "Global"),
                    Direction = GetArrayValue(directions, i),
                    Distance1 = GetArrayValue(dist1, i),
                    Distance2 = GetArrayValue(dist2, i),
                    Value1 = GetArrayValue(val1, i),
                    Value2 = GetArrayValue(val2, i),
                    IsRelativeDistance = GetArrayValue(rd1, i) >= 0 || GetArrayValue(rd2, i) >= 0
                });
            }

            return OperationResult<IReadOnlyList<FrameLoadInfo>>.Success(loads);
        }

        internal static OperationResult<IReadOnlyList<FrameLoadInfo>> GetPointLoads<TSapModel>(
            string productName,
            TSapModel sapModel,
            string frameName,
            CSISapModelReadFramePointLoads<TSapModel> getLoadPoint)
        {
            if (string.IsNullOrWhiteSpace(frameName))
            {
                return OperationResult<IReadOnlyList<FrameLoadInfo>>.Failure("Frame name is required.");
            }

            int numberItems = 0;
            string[] frameNames = null;
            string[] loadPatterns = null;
            int[] loadTypes = null;
            string[] coordinateSystems = null;
            int[] directions = null;
            double[] relativeDistance = null;
            double[] distance = null;
            double[] value = null;

            int result = getLoadPoint(sapModel, frameName, ref numberItems, ref frameNames, ref loadPatterns, ref loadTypes, ref coordinateSystems, ref directions, ref relativeDistance, ref distance, ref value);
            if (result != 0)
            {
                return OperationResult<IReadOnlyList<FrameLoadInfo>>.Failure($"{productName} FrameObj.GetLoadPoint failed for '{frameName}' (return code {result}).");
            }

            var loads = new List<FrameLoadInfo>();
            for (int i = 0; i < numberItems; i++)
            {
                loads.Add(new FrameLoadInfo
                {
                    FrameName = GetArrayValue(frameNames, i, frameName),
                    LoadPattern = GetArrayValue(loadPatterns, i, string.Empty),
                    LoadType = "Point",
                    CoordinateSystem = GetArrayValue(coordinateSystems, i, "Global"),
                    Direction = GetArrayValue(directions, i),
                    Distance1 = GetArrayValue(distance, i),
                    Distance2 = GetArrayValue(relativeDistance, i),
                    Value1 = GetArrayValue(value, i),
                    IsRelativeDistance = GetArrayValue(relativeDistance, i) >= 0
                });
            }

            return OperationResult<IReadOnlyList<FrameLoadInfo>>.Success(loads);
        }

        private static OperationResult<CSISapModelAddFramesResultDTO> AddFrames<TSapModel, TFrameInput, TAddResult>(
            IReadOnlyList<TFrameInput> frameInputs,
            string progressTitle,
            string productName,
            string apiCallName,
            string operationName,
            TSapModel sapModel,
            Func<TSapModel, OperationResult> refreshView,
            Func<TSapModel, TFrameInput, TAddResult> addFrame,
            Func<TAddResult, string> getWarningMessage)
            where TFrameInput : class
        {
            if (frameInputs == null || frameInputs.Count == 0)
            {
                return OperationResult<CSISapModelAddFramesResultDTO>.Failure("No valid rows were found in the selected range.");
            }

            try
            {
                var failedRowMessages = new List<string>();
                var successCount = 0;

                BatchProgressWindow.RunWithProgress(frameInputs.Count, progressTitle, ctx =>
                {
                    foreach (var frameInput in frameInputs)
                    {
                        if (ctx.IsCancellationRequested) break;

                        var result = addFrame(sapModel, frameInput);
                        int returnCode = GetFrameAddReturnCode(result);
                        int excelRowNumber = GetExcelRowNumber(frameInput);

                        if (returnCode != 0)
                        {
                            failedRowMessages.Add($"Row {excelRowNumber}: {productName} API call {apiCallName} failed (return code {returnCode}).");
                            ctx.IncrementSkipped();
                            continue;
                        }

                        successCount++;
                        ctx.IncrementRan();

                        var warningMessage = getWarningMessage == null ? null : getWarningMessage(result);
                        if (!string.IsNullOrWhiteSpace(warningMessage))
                        {
                            failedRowMessages.Add($"Row {excelRowNumber}: {warningMessage}");
                        }
                    }
                });

                if (successCount > 0)
                {
                    var refreshResult = refreshView(sapModel);
                    if (!refreshResult.IsSuccess)
                    {
                        return OperationResult<CSISapModelAddFramesResultDTO>.Failure(refreshResult.Message);
                    }
                }

                return OperationResult<CSISapModelAddFramesResultDTO>.Success(new CSISapModelAddFramesResultDTO
                {
                    AddedCount = successCount,
                    FailedRowMessages = failedRowMessages
                });
            }
            catch (COMException ex)
            {
                return OperationResult<CSISapModelAddFramesResultDTO>.Failure($"{productName} COM error while adding frames by {operationName}: {ex.Message}");
            }
            catch (Exception ex)
            {
                return OperationResult<CSISapModelAddFramesResultDTO>.Failure($"{productName} add-by-{operationName} failed unexpectedly: {ex.Message}");
            }
        }

        private static OperationResult SelectObjectsByUniqueNames<TSapModel>(
            IReadOnlyList<string> uniqueNames,
            string objectTypeName,
            string productName,
            TSapModel sapModel,
            CSISapModelClearSelection<TSapModel> clearSelection,
            CSISapModelSetSelectedByName<TSapModel> setSelected,
            Func<TSapModel, OperationResult> refreshView)
        {
            if (uniqueNames == null || uniqueNames.Count == 0)
            {
                return OperationResult.Failure("The selected Excel range does not contain any non-empty values.");
            }

            var orderedUniqueNames = GetOrderedDistinctNames(uniqueNames);
            if (orderedUniqueNames.Count == 0)
            {
                return OperationResult.Failure("The selected Excel range does not contain any non-empty values.");
            }

            try
            {
                int clearSelectionResult = clearSelection(sapModel);
                if (clearSelectionResult != 0)
                {
                    return OperationResult.Failure($"Failed to clear {productName} selection before selecting {objectTypeName}s by UniqueName.");
                }

                var unresolved = new List<string>();
                var selectedCount = 0;

                BatchProgressWindow.RunWithProgress(orderedUniqueNames.Count, $"Selecting {objectTypeName}s...", ctx =>
                {
                    foreach (var uniqueName in orderedUniqueNames)
                    {
                        if (ctx.IsCancellationRequested) break;

                        int result = setSelected(sapModel, uniqueName);
                        if (result == 0)
                        {
                            selectedCount++;
                            ctx.IncrementRan();
                        }
                        else
                        {
                            unresolved.Add(uniqueName);
                            ctx.IncrementSkipped();
                        }
                    }
                });

                var message = $"Selected {selectedCount} {objectTypeName}(s) by UniqueName.";
                if (unresolved.Count > 0)
                {
                    message += $" Not found: {string.Join(", ", unresolved)}.";
                }

                var refreshResult = refreshView(sapModel);
                if (!refreshResult.IsSuccess)
                {
                    return refreshResult;
                }

                return OperationResult.Success(message);
            }
            catch
            {
                return OperationResult.Failure($"Failed to select {productName} {objectTypeName}s by UniqueName.");
            }
        }

        private static int GetFrameAddReturnCode<TAddResult>(TAddResult result)
        {
            if (result is int returnCode)
            {
                return returnCode;
            }

            if (result is FrameAddResult frameAddResult)
            {
                return frameAddResult.ReturnCode;
            }

            throw new InvalidOperationException("Unsupported frame add result.");
        }

        private static int GetExcelRowNumber<TFrameInput>(TFrameInput frameInput)
        {
            var property = frameInput.GetType().GetProperty("ExcelRowNumber");
            return property == null ? 0 : (int)property.GetValue(frameInput, null);
        }

        private static FrameAddResultDto AddOneFrameObject<TSapModel>(
            FrameAddRequestDto request,
            string productName,
            TSapModel sapModel,
            CSISapModelAddFrameByPoint<TSapModel> addByPoint,
            CSISapModelAddFrameByCoord<TSapModel> addByCoord)
        {
            if (request == null)
            {
                return Failed(null, null, "Frame definition is required.", null);
            }

            string userName = FirstCleanName(request.UserName, request.UniqueName, request.FrameName, request.Name);
            string propName = FirstCleanName(request.PropName, request.SectionName);
            if (string.IsNullOrWhiteSpace(propName))
            {
                propName = "Default";
            }

            string pointIName = CleanName(request.PointIName);
            string pointJName = CleanName(request.PointJName);
            if (!string.IsNullOrWhiteSpace(pointIName) && !string.IsNullOrWhiteSpace(pointJName))
            {
                string createdName = string.Empty;
                int result = addByPoint(
                    sapModel,
                    new CSISapModelFrameByPointInput
                    {
                        Point1Name = pointIName,
                        Point2Name = pointJName,
                        SectionName = propName,
                        UniqueName = userName
                    },
                    ref createdName,
                    propName,
                    userName);

                return result == 0
                    ? Succeeded("AddByPoint", createdName, userName)
                    : Failed("AddByPoint", createdName, $"{productName} FrameObj.AddByPoint failed (return code {result}).", result);
            }

            if (HasAllCoordinates(request))
            {
                string createdName = string.Empty;
                int result = addByCoord(
                    sapModel,
                    new CSISapModelFrameByCoordInput
                    {
                        Xi = request.Xi.Value,
                        Yi = request.Yi.Value,
                        Zi = request.Zi.Value,
                        Xj = request.Xj.Value,
                        Yj = request.Yj.Value,
                        Zj = request.Zj.Value,
                        SectionName = propName,
                        UniqueName = userName
                    },
                    ref createdName,
                    propName,
                    userName);

                return result == 0
                    ? Succeeded("AddByCoord", createdName, userName)
                    : Failed("AddByCoord", createdName, $"{productName} FrameObj.AddByCoord failed (return code {result}).", result);
            }

            return Failed(null, null, "Provide PointIName and PointJName, or all six coordinates Xi, Yi, Zi, Xj, Yj, Zj.", null);
        }

        private static bool HasAllCoordinates(FrameAddRequestDto request)
        {
            return request.Xi.HasValue &&
                   request.Yi.HasValue &&
                   request.Zi.HasValue &&
                   request.Xj.HasValue &&
                   request.Yj.HasValue &&
                   request.Zj.HasValue;
        }

        private static FrameAddResultDto Succeeded(string addMethod, string createdName, string userName)
        {
            string frameName = CleanName(createdName);
            if (string.IsNullOrWhiteSpace(frameName))
            {
                frameName = CleanName(userName);
            }

            return new FrameAddResultDto
            {
                Success = true,
                FrameName = frameName,
                AddMethod = addMethod,
                FailureReason = null,
                ReturnCode = 0
            };
        }

        private static FrameAddResultDto Failed(string addMethod, string frameName, string failureReason, int? returnCode)
        {
            return new FrameAddResultDto
            {
                Success = false,
                FrameName = CleanName(frameName),
                AddMethod = addMethod,
                FailureReason = failureReason,
                ReturnCode = returnCode
            };
        }

        private static IReadOnlyList<string> GetOrderedDistinctNames(IReadOnlyList<string> names)
        {
            var uniqueNames = new List<string>();
            var seenNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (var rawName in names)
            {
                var name = string.IsNullOrWhiteSpace(rawName) ? null : rawName.Trim();
                if (string.IsNullOrWhiteSpace(name) || seenNames.Contains(name))
                {
                    continue;
                }

                seenNames.Add(name);
                uniqueNames.Add(name);
            }

            return uniqueNames;
        }

        private static string CleanName(string value)
        {
            return string.IsNullOrWhiteSpace(value) ? null : value.Trim();
        }

        private static string FirstCleanName(params string[] values)
        {
            if (values == null)
            {
                return null;
            }

            foreach (string value in values)
            {
                string clean = CleanName(value);
                if (!string.IsNullOrWhiteSpace(clean))
                {
                    return clean;
                }
            }

            return null;
        }

        private static string GetArrayValue(string[] values, int index, string fallback)
        {
            return values == null || index < 0 || index >= values.Length || string.IsNullOrWhiteSpace(values[index])
                ? fallback
                : values[index];
        }

        private static int GetArrayValue(int[] values, int index)
        {
            return values == null || index < 0 || index >= values.Length ? 0 : values[index];
        }

        private static double GetArrayValue(double[] values, int index)
        {
            return values == null || index < 0 || index >= values.Length ? 0 : values[index];
        }

        private class FrameAddResult
        {
            internal FrameAddResult(int returnCode, string warningMessage)
            {
                ReturnCode = returnCode;
                WarningMessage = warningMessage;
            }

            internal int ReturnCode { get; }

            internal string WarningMessage { get; }
        }
    }
}


