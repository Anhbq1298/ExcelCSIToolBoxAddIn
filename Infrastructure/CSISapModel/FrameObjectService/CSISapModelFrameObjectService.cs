using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using ExcelCSIToolBoxAddIn.Common.Results;
using ExcelCSIToolBoxAddIn.UI.Views;

namespace ExcelCSIToolBoxAddIn.Infrastructure.CSISapModel
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

    internal static class CSISapModelFrameObjectService
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

        internal static OperationResult<CSISapModelAddFramesResult> AddFramesByCoordinates<TSapModel>(
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

        internal static OperationResult<CSISapModelAddFramesResult> AddFramesByPoint<TSapModel>(
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

        internal static OperationResult<IReadOnlyList<string>> GetSelectedFramesFromActiveModel<TSapModel>(
            string productName,
            TSapModel sapModel,
            CSISapModelReadSelectedObjects<TSapModel> getSelectedObjects)
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

        private static OperationResult<CSISapModelAddFramesResult> AddFrames<TSapModel, TFrameInput, TAddResult>(
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
                return OperationResult<CSISapModelAddFramesResult>.Failure("No valid rows were found in the selected range.");
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
                        return OperationResult<CSISapModelAddFramesResult>.Failure(refreshResult.Message);
                    }
                }

                return OperationResult<CSISapModelAddFramesResult>.Success(new CSISapModelAddFramesResult
                {
                    AddedCount = successCount,
                    FailedRowMessages = failedRowMessages
                });
            }
            catch (COMException ex)
            {
                return OperationResult<CSISapModelAddFramesResult>.Failure($"{productName} COM error while adding frames by {operationName}: {ex.Message}");
            }
            catch (Exception ex)
            {
                return OperationResult<CSISapModelAddFramesResult>.Failure($"{productName} add-by-{operationName} failed unexpectedly: {ex.Message}");
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
