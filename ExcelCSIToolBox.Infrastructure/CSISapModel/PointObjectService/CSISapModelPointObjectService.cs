using System.Runtime.InteropServices;
using System;
using System.Collections.Generic;
using ExcelCSIToolBox.Core.Common.Results;
using ExcelCSIToolBox.Data;
using ExcelCSIToolBox.Data.CSISapModel.PointObject;
using ExcelCSIToolBox.Data.DTOs.CSI;
using ExcelCSIToolBox.Data.Models;

namespace ExcelCSIToolBox.Infrastructure.CSISapModel
{
    public delegate int CSISapModelAddCartesianPoint<TSapModel>(
        TSapModel sapModel,
        CSISapModelPointCartesianInput pointInput,
        ref string assignedName,
        string requestedUniqueName);

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

    internal delegate int CSISapModelReadPointSelected<TSapModel>(
        TSapModel sapModel,
        string pointName,
        ref bool selected);

    internal delegate int CSISapModelReadPointRestraint<TSapModel>(
        TSapModel sapModel,
        string pointName,
        ref bool[] restraints);

    internal delegate int CSISapModelReadPointLoadForce<TSapModel>(
        TSapModel sapModel,
        string pointName,
        ref int numberItems,
        ref string[] pointNames,
        ref string[] loadPatterns,
        ref int[] caseSteps,
        ref string[] coordinateSystems,
        ref double[] f1,
        ref double[] f2,
        ref double[] f3,
        ref double[] m1,
        ref double[] m2,
        ref double[] m3);

    internal delegate int CSISapModelSetPointRestraint<TSapModel>(
        TSapModel sapModel,
        string pointName,
        ref bool[] restraints);

    internal delegate int CSISapModelSetPointLoadForce<TSapModel>(
        TSapModel sapModel,
        string pointName,
        string loadPattern,
        ref double[] forceValues,
        bool replace,
        string coordinateSystem);

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

        internal static OperationResult<IReadOnlyList<string>> GetNameList<TSapModel>(
            string productName,
            TSapModel sapModel,
            CSISapModelGetPointNames<TSapModel> getNameList)
        {
            int numberNames = 0;
            string[] names = null;
            int result = getNameList(sapModel, ref numberNames, ref names);
            if (result != 0)
            {
                return OperationResult<IReadOnlyList<string>>.Failure($"{productName} PointObj.GetNameList failed (return code {result}).");
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
                return OperationResult<int>.Failure($"{productName} PointObj.Count failed (return code {result}).");
            }

            return OperationResult<int>.Success(count);
        }

        internal static OperationResult<PointObjectInfo> GetByName<TSapModel>(
            string productName,
            TSapModel sapModel,
            string pointName,
            CSISapModelReadPointCoordinates<TSapModel> getPointCoordinates,
            CSISapModelReadPointSelected<TSapModel> getSelected)
        {
            if (string.IsNullOrWhiteSpace(pointName))
            {
                return OperationResult<PointObjectInfo>.Failure("Point name is required.");
            }

            double x = 0;
            double y = 0;
            double z = 0;
            int coordResult = getPointCoordinates(sapModel, pointName, ref x, ref y, ref z);
            if (coordResult != 0)
            {
                return OperationResult<PointObjectInfo>.Failure($"{productName} point '{pointName}' was not found (return code {coordResult}).");
            }

            bool selected = false;
            if (getSelected != null)
            {
                getSelected(sapModel, pointName, ref selected);
            }

            return OperationResult<PointObjectInfo>.Success(new PointObjectInfo
            {
                Name = pointName,
                X = x,
                Y = y,
                Z = z,
                IsSelected = selected,
                CoordinateSystem = "Global"
            });
        }

        internal static OperationResult<PointRestraintInfo> GetRestraint<TSapModel>(
            string productName,
            TSapModel sapModel,
            string pointName,
            CSISapModelReadPointRestraint<TSapModel> getRestraint)
        {
            if (string.IsNullOrWhiteSpace(pointName))
            {
                return OperationResult<PointRestraintInfo>.Failure("Point name is required.");
            }

            bool[] values = null;
            int result = getRestraint(sapModel, pointName, ref values);
            if (result != 0 || values == null || values.Length < 6)
            {
                return OperationResult<PointRestraintInfo>.Failure($"{productName} PointObj.GetRestraint failed for '{pointName}' (return code {result}).");
            }

            return OperationResult<PointRestraintInfo>.Success(new PointRestraintInfo
            {
                PointName = pointName,
                U1 = values[0],
                U2 = values[1],
                U3 = values[2],
                R1 = values[3],
                R2 = values[4],
                R3 = values[5]
            });
        }

        internal static OperationResult<IReadOnlyList<PointLoadInfo>> GetLoadForces<TSapModel>(
            string productName,
            TSapModel sapModel,
            string pointName,
            CSISapModelReadPointLoadForce<TSapModel> getLoadForce)
        {
            if (string.IsNullOrWhiteSpace(pointName))
            {
                return OperationResult<IReadOnlyList<PointLoadInfo>>.Failure("Point name is required.");
            }

            int numberItems = 0;
            string[] pointNames = null;
            string[] loadPatterns = null;
            int[] caseSteps = null;
            string[] coordinateSystems = null;
            double[] f1 = null;
            double[] f2 = null;
            double[] f3 = null;
            double[] m1 = null;
            double[] m2 = null;
            double[] m3 = null;

            int result = getLoadForce(sapModel, pointName, ref numberItems, ref pointNames, ref loadPatterns, ref caseSteps, ref coordinateSystems, ref f1, ref f2, ref f3, ref m1, ref m2, ref m3);
            if (result != 0)
            {
                return OperationResult<IReadOnlyList<PointLoadInfo>>.Failure($"{productName} PointObj.GetLoadForce failed for '{pointName}' (return code {result}).");
            }

            var loads = new List<PointLoadInfo>();
            for (int i = 0; i < numberItems; i++)
            {
                loads.Add(new PointLoadInfo
                {
                    PointName = GetArrayValue(pointNames, i, pointName),
                    LoadPattern = GetArrayValue(loadPatterns, i, string.Empty),
                    CoordinateSystem = GetArrayValue(coordinateSystems, i, "Global"),
                    F1 = GetArrayValue(f1, i),
                    F2 = GetArrayValue(f2, i),
                    F3 = GetArrayValue(f3, i),
                    M1 = GetArrayValue(m1, i),
                    M2 = GetArrayValue(m2, i),
                    M3 = GetArrayValue(m3, i)
                });
            }

            return OperationResult<IReadOnlyList<PointLoadInfo>>.Success(loads);
        }

        internal static OperationResult SetRestraint<TSapModel>(
            string productName,
            TSapModel sapModel,
            IReadOnlyList<string> pointNames,
            IReadOnlyList<bool> restraints,
            CSISapModelSetPointRestraint<TSapModel> setRestraint,
            Func<TSapModel, OperationResult> refreshView)
        {
            if (pointNames == null || pointNames.Count == 0)
            {
                return OperationResult.Failure("At least one point name is required.");
            }

            if (restraints == null || restraints.Count != 6)
            {
                return OperationResult.Failure("Point restraints must contain exactly 6 boolean values.");
            }

            int success = 0;
            var failures = new List<string>();
            foreach (string rawName in pointNames)
            {
                string pointName = CleanName(rawName);
                if (string.IsNullOrWhiteSpace(pointName))
                {
                    continue;
                }

                bool[] values = ToArray(restraints);
                int result = setRestraint(sapModel, pointName, ref values);
                if (result == 0)
                {
                    success++;
                }
                else
                {
                    failures.Add($"{pointName}: return code {result}");
                }
            }

            refreshView(sapModel);
            string message = $"Set restraints for {success} point object(s).";
            if (failures.Count > 0)
            {
                message += " Failed: " + string.Join("; ", failures);
            }

            return failures.Count == 0 ? OperationResult.Success(message) : OperationResult.Failure(message);
        }

        internal static OperationResult SetLoadForce<TSapModel>(
            string productName,
            TSapModel sapModel,
            IReadOnlyList<string> pointNames,
            string loadPattern,
            IReadOnlyList<double> forceValues,
            bool replace,
            string coordinateSystem,
            CSISapModelSetPointLoadForce<TSapModel> setLoadForce,
            Func<TSapModel, OperationResult> refreshView)
        {
            if (pointNames == null || pointNames.Count == 0)
            {
                return OperationResult.Failure("At least one point name is required.");
            }

            if (string.IsNullOrWhiteSpace(loadPattern))
            {
                return OperationResult.Failure("Load pattern is required.");
            }

            if (forceValues == null || forceValues.Count != 6)
            {
                return OperationResult.Failure("Point load force values must contain exactly 6 numbers.");
            }

            string cSys = string.IsNullOrWhiteSpace(coordinateSystem) ? "Global" : coordinateSystem.Trim();
            int success = 0;
            var failures = new List<string>();
            foreach (string rawName in pointNames)
            {
                string pointName = CleanName(rawName);
                if (string.IsNullOrWhiteSpace(pointName))
                {
                    continue;
                }

                double[] values = ToArray(forceValues);
                int result = setLoadForce(sapModel, pointName, loadPattern.Trim(), ref values, replace, cSys);
                if (result == 0)
                {
                    success++;
                }
                else
                {
                    failures.Add($"{pointName}: return code {result}");
                }
            }

            refreshView(sapModel);
            string message = $"Assigned point load '{loadPattern}' to {success} point object(s).";
            if (failures.Count > 0)
            {
                message += " Failed: " + string.Join("; ", failures);
            }

            return failures.Count == 0 ? OperationResult.Success(message) : OperationResult.Failure(message);
        }

        internal static OperationResult<IReadOnlyList<CSISapModelPointDataDTO>> GetSelectedPointsFromActiveModel<TSapModel>(
            string productName,
            TSapModel sapModel,
            CSISapModelGetSelectedObjects<TSapModel> getSelectedObjects,
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

        private static string CleanName(string value)
        {
            return string.IsNullOrWhiteSpace(value) ? null : value.Trim();
        }

        private static bool[] ToArray(IReadOnlyList<bool> values)
        {
            var result = new bool[values.Count];
            for (int i = 0; i < values.Count; i++)
            {
                result[i] = values[i];
            }

            return result;
        }

        private static double[] ToArray(IReadOnlyList<double> values)
        {
            var result = new double[values.Count];
            for (int i = 0; i < values.Count; i++)
            {
                result[i] = values[i];
            }

            return result;
        }

        private static string GetArrayValue(string[] values, int index, string fallback)
        {
            return values == null || index < 0 || index >= values.Length || string.IsNullOrWhiteSpace(values[index])
                ? fallback
                : values[index];
        }

        private static double GetArrayValue(double[] values, int index)
        {
            return values == null || index < 0 || index >= values.Length ? 0 : values[index];
        }
    }
}


