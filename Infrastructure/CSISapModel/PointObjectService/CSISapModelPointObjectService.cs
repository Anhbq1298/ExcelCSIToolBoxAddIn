using System.Collections.Generic;
using ExcelCSIToolBoxAddIn.Common.Results;
using ExcelCSIToolBoxAddIn.UI.Views;

namespace ExcelCSIToolBoxAddIn.Infrastructure.CSISapModel
{
    public delegate int CSISapModelAddCartesianPoint<TSapModel>(
        TSapModel sapModel,
        CSISapModelPointCartesianInput pointInput,
        ref string assignedName,
        string requestedUniqueName);

    internal delegate int CSISapModelClearSelection<TSapModel>(TSapModel sapModel);

    internal delegate int CSISapModelSetSelectedByName<TSapModel>(
        TSapModel sapModel,
        string objectName);

    internal delegate int CSISapModelReadSelectedObjects<TSapModel>(
        TSapModel sapModel,
        ref int numberItems,
        ref int[] objectTypes,
        ref string[] objectNames);

    internal delegate int CSISapModelReadPointCoordinates<TSapModel>(
        TSapModel sapModel,
        string pointName,
        ref double x,
        ref double y,
        ref double z);

    internal delegate int CSISapModelReadPointLabel<TSapModel>(
        TSapModel sapModel,
        string pointName,
        ref string pointLabel,
        ref string pointStory);

    internal static class CSISapModelPointObjectService
    {
        internal static OperationResult SelectPointsByUniqueNames<TSapModel>(
            IReadOnlyList<string> uniqueNames,
            string productName,
            TSapModel sapModel,
            CSISapModelClearSelection<TSapModel> clearSelection,
            CSISapModelSetSelectedByName<TSapModel> setSelected,
            Func<TSapModel, OperationResult> refreshView)
        {
            return SelectObjectsByUniqueNames(
                uniqueNames,
                "point",
                productName,
                sapModel,
                clearSelection,
                setSelected,
                refreshView);
        }

        internal static OperationResult<CSISapModelAddPointsResultDTO> AddPointsByCartesian<TSapModel>(
            IReadOnlyList<CSISapModelPointCartesianInput> pointInputs,
            string productName,
            TSapModel sapModel,
            CSISapModelAddCartesianPoint<TSapModel> addPoint,
            Func<TSapModel, OperationResult> refreshView)
        {
            if (pointInputs == null || pointInputs.Count == 0)
            {
                return OperationResult<CSISapModelAddPointsResultDTO>.Failure("No valid rows were found in the selected range.");
            }

            try
            {
                var failedRowMessages = new List<string>();
                var successCount = 0;

                BatchProgressWindow.RunWithProgress(pointInputs.Count, "Adding Points to Model...", ctx =>
                {
                    foreach (var pointInput in pointInputs)
                    {
                        if (ctx.IsCancellationRequested) break;

                        string assignedName = string.Empty;
                        string requestedUniqueName = string.IsNullOrWhiteSpace(pointInput.UniqueName) ? string.Empty : pointInput.UniqueName;
                        int addResult = addPoint(sapModel, pointInput, ref assignedName, requestedUniqueName);

                        if (addResult != 0)
                        {
                            failedRowMessages.Add($"Row {pointInput.ExcelRowNumber}: {productName} API call PointObj.AddCartesian failed (return code {addResult}).");
                            ctx.IncrementSkipped();
                            continue;
                        }

                        successCount++;
                        ctx.IncrementRan();

                        if (!string.IsNullOrWhiteSpace(requestedUniqueName) &&
                            !string.Equals(assignedName, requestedUniqueName, StringComparison.OrdinalIgnoreCase))
                        {
                            failedRowMessages.Add($"Row {pointInput.ExcelRowNumber}: Point was created, but {productName} assigned UniqueName '{assignedName}' instead of requested '{requestedUniqueName}'.");
                        }
                    }
                });

                if (successCount > 0)
                {
                    var refreshResult = refreshView(sapModel);
                    if (!refreshResult.IsSuccess)
                    {
                        return OperationResult<CSISapModelAddPointsResultDTO>.Failure(refreshResult.Message);
                    }
                }

                return OperationResult<CSISapModelAddPointsResultDTO>.Success(new CSISapModelAddPointsResultDTO
                {
                    AddedCount = successCount,
                    FailedRowMessages = failedRowMessages
                });
            }
            catch (COMException ex)
            {
                return OperationResult<CSISapModelAddPointsResultDTO>.Failure($"{productName} COM error while adding points by Cartesian coordinates: {ex.Message}");
            }
            catch (Exception ex)
            {
                return OperationResult<CSISapModelAddPointsResultDTO>.Failure($"{productName} add-by-Cartesian failed unexpectedly: {ex.Message}");
            }
        }

        internal static OperationResult<IReadOnlyList<CSISapModelPointDataDTO>> GetSelectedPointsFromActiveModel<TSapModel>(
            string productName,
            TSapModel sapModel,
            CSISapModelReadSelectedObjects<TSapModel> getSelectedObjects,
            CSISapModelReadPointCoordinates<TSapModel> getPointCoordinates,
            CSISapModelReadPointLabel<TSapModel> getPointLabel)
        {
            try
            {
                int numberItems = 0;
                int[] objectTypes = null;
                string[] objectNames = null;
                int selectedResult = getSelectedObjects(sapModel, ref numberItems, ref objectTypes, ref objectNames);

                if (selectedResult != 0)
                {
                    return OperationResult<IReadOnlyList<CSISapModelPointDataDTO>>.Failure($"Failed to read selected objects from {productName}.");
                }

                var points = new List<CSISapModelPointDataDTO>();

                for (int i = 0; i < numberItems; i++)
                {
                    if (objectTypes == null || objectNames == null || i >= objectTypes.Length || i >= objectNames.Length)
                    {
                        continue;
                    }

                    if (objectTypes[i] != CSISapModelObjectTypeIds.Point || string.IsNullOrWhiteSpace(objectNames[i]))
                    {
                        continue;
                    }

                    double x = 0;
                    double y = 0;
                    double z = 0;
                    int pointResult = getPointCoordinates(sapModel, objectNames[i], ref x, ref y, ref z);
                    if (pointResult != 0)
                    {
                        continue;
                    }

                    string pointLabel = string.Empty;
                    string pointStory = string.Empty;
                    int pointLabelResult = getPointLabel == null
                        ? -1
                        : getPointLabel(sapModel, objectNames[i], ref pointLabel, ref pointStory);

                    points.Add(new CSISapModelPointDataDTO
                    {
                        PointUniqueName = objectNames[i],
                        PointLabel = pointLabelResult == 0 && !string.IsNullOrWhiteSpace(pointLabel)
                            ? pointLabel
                            : objectNames[i],
                        X = x,
                        Y = y,
                        Z = z
                    });
                }

                if (points.Count == 0)
                {
                    return OperationResult<IReadOnlyList<CSISapModelPointDataDTO>>.Failure($"No point objects are selected in {productName}.");
                }

                return OperationResult<IReadOnlyList<CSISapModelPointDataDTO>>.Success(points);
            }
            catch
            {
                return OperationResult<IReadOnlyList<CSISapModelPointDataDTO>>.Failure($"Unable to read selected {productName} points.");
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
    }
}
