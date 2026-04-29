using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using ExcelCSIToolBox.Core.Common.Results;
using ExcelCSIToolBox.Core.Geometry;
using ExcelCSIToolBox.Core.Models.CSI;
using ExcelCSIToolBox.Data;
using ExcelCSIToolBox.Data.DTOs.CSI;
using ExcelCSIToolBox.Data.Models;

namespace ExcelCSIToolBox.Infrastructure.CSISapModel
{
    internal delegate int CSISapModelGetFramePoints<TSapModel>(
        TSapModel sapModel,
        string frameName,
        ref string point1Name,
        ref string point2Name);

    internal delegate int CSISapModelAddAreaByCoordinates<TSapModel>(
        TSapModel sapModel,
        int nodeCount,
        ref double[] x,
        ref double[] y,
        ref double[] z,
        ref string areaName,
        string propertyName);

    internal delegate int CSISapModelAreaGetNameList<TSapModel>(
        TSapModel sapModel,
        ref int numberNames,
        ref string[] names);

    internal delegate int CSISapModelAreaGetPoints<TSapModel>(
        TSapModel sapModel,
        string areaName,
        ref int numberPoints,
        ref string[] pointNames);

    internal delegate int CSISapModelAreaGetProperty<TSapModel>(
        TSapModel sapModel,
        string areaName,
        ref string propertyName);

    internal delegate int CSISapModelAreaGetSelected<TSapModel>(
        TSapModel sapModel,
        string areaName,
        ref bool selected);

    internal delegate int CSISapModelAreaGetLoadUniform<TSapModel>(
        TSapModel sapModel,
        string areaName,
        ref int numberItems,
        ref string[] areaNames,
        ref string[] loadPatterns,
        ref string[] coordinateSystems,
        ref int[] directions,
        ref double[] values);

    internal delegate int CSISapModelAreaAddByPoint<TSapModel>(
        TSapModel sapModel,
        int numberPoints,
        ref string[] pointNames,
        ref string areaName,
        string propertyName,
        string userName);

    internal delegate int CSISapModelAreaAddByCoord<TSapModel>(
        TSapModel sapModel,
        int numberPoints,
        ref double[] x,
        ref double[] y,
        ref double[] z,
        ref string areaName,
        string propertyName,
        string userName,
        string coordinateSystem);

    internal delegate int CSISapModelAreaSetLoadUniform<TSapModel>(
        TSapModel sapModel,
        string areaName,
        string loadPattern,
        double value,
        int direction,
        bool replace,
        string coordinateSystem);

    internal delegate int CSISapModelAreaDelete<TSapModel>(
        TSapModel sapModel,
        string areaName);

    internal static class CSISapModelShellObjectService
    {
        internal static OperationResult<IReadOnlyList<string>> GetNameList<TSapModel>(
            TSapModel sapModel,
            string productName,
            CSISapModelAreaGetNameList<TSapModel> getNameList)
        {
            try
            {
                int numberNames = 0;
                string[] names = null;
                int result = getNameList(sapModel, ref numberNames, ref names);
                if (result != 0)
                {
                    return OperationResult<IReadOnlyList<string>>.Failure($"{productName} AreaObj.GetNameList failed (return code {result}).");
                }

                return OperationResult<IReadOnlyList<string>>.Success((names ?? new string[0]).Take(numberNames).ToList());
            }
            catch (COMException ex)
            {
                return OperationResult<IReadOnlyList<string>>.Failure($"{productName} COM error while reading shell/area names: {ex.Message}");
            }
            catch (Exception ex)
            {
                return OperationResult<IReadOnlyList<string>>.Failure($"Failed to read shell/area names: {ex.Message}");
            }
        }

        internal static OperationResult<int> GetCount<TSapModel>(
            TSapModel sapModel,
            string productName,
            CSISapModelReadCount<TSapModel> getCount)
        {
            try
            {
                int count = 0;
                int result = getCount(sapModel, ref count);
                if (result != 0)
                {
                    return OperationResult<int>.Failure($"{productName} AreaObj.Count failed (return code {result}).");
                }

                return OperationResult<int>.Success(count);
            }
            catch (COMException ex)
            {
                return OperationResult<int>.Failure($"{productName} COM error while reading shell/area count: {ex.Message}");
            }
            catch (Exception ex)
            {
                return OperationResult<int>.Failure($"Failed to read shell/area count: {ex.Message}");
            }
        }

        internal static OperationResult<IReadOnlyList<string>> GetPoints<TSapModel>(
            TSapModel sapModel,
            string productName,
            string areaName,
            CSISapModelAreaGetPoints<TSapModel> getPoints)
        {
            var validation = ShellObjectValidation.ValidateAreaName(areaName);
            if (!validation.IsSuccess)
            {
                return OperationResult<IReadOnlyList<string>>.Failure(validation.Message);
            }

            try
            {
                int numberPoints = 0;
                string[] pointNames = null;
                int result = getPoints(sapModel, areaName, ref numberPoints, ref pointNames);
                if (result != 0)
                {
                    return OperationResult<IReadOnlyList<string>>.Failure($"{productName} AreaObj.GetPoints failed for '{areaName}' (return code {result}).");
                }

                return OperationResult<IReadOnlyList<string>>.Success((pointNames ?? new string[0]).Take(numberPoints).ToList());
            }
            catch (COMException ex)
            {
                return OperationResult<IReadOnlyList<string>>.Failure($"{productName} COM error while reading shell/area points: {ex.Message}");
            }
            catch (Exception ex)
            {
                return OperationResult<IReadOnlyList<string>>.Failure($"Failed to read shell/area points: {ex.Message}");
            }
        }

        internal static OperationResult<string> GetProperty<TSapModel>(
            TSapModel sapModel,
            string productName,
            string areaName,
            CSISapModelAreaGetProperty<TSapModel> getProperty)
        {
            var validation = ShellObjectValidation.ValidateAreaName(areaName);
            if (!validation.IsSuccess)
            {
                return OperationResult<string>.Failure(validation.Message);
            }

            try
            {
                string propertyName = string.Empty;
                int result = getProperty(sapModel, areaName, ref propertyName);
                if (result != 0)
                {
                    return OperationResult<string>.Failure($"{productName} AreaObj.GetProperty failed for '{areaName}' (return code {result}).");
                }

                return OperationResult<string>.Success(propertyName);
            }
            catch (COMException ex)
            {
                return OperationResult<string>.Failure($"{productName} COM error while reading shell/area property: {ex.Message}");
            }
            catch (Exception ex)
            {
                return OperationResult<string>.Failure($"Failed to read shell/area property: {ex.Message}");
            }
        }

        internal static OperationResult<IReadOnlyList<string>> GetSelectedShells<TSapModel>(
            TSapModel sapModel,
            string productName,
            CSISapModelGetSelectedObjects<TSapModel> getSelectedObjects)
        {
            try
            {
                int numberItems = 0;
                int[] objectTypes = null;
                string[] objectNames = null;
                int result = getSelectedObjects(sapModel, ref numberItems, ref objectTypes, ref objectNames);
                if (result != 0)
                {
                    return OperationResult<IReadOnlyList<string>>.Failure($"{productName} SelectObj.GetSelected failed (return code {result}).");
                }

                var areaNames = new List<string>();
                for (int i = 0; i < numberItems; i++)
                {
                    if (objectTypes == null || objectNames == null || i >= objectTypes.Length || i >= objectNames.Length)
                    {
                        continue;
                    }

                    if (objectTypes[i] == CSISapModelObjectTypeIds.Shell && !string.IsNullOrWhiteSpace(objectNames[i]))
                    {
                        areaNames.Add(objectNames[i]);
                    }
                }

                return OperationResult<IReadOnlyList<string>>.Success(areaNames);
            }
            catch (COMException ex)
            {
                return OperationResult<IReadOnlyList<string>>.Failure($"{productName} COM error while reading selected shell/area objects: {ex.Message}");
            }
            catch (Exception ex)
            {
                return OperationResult<IReadOnlyList<string>>.Failure($"Failed to read selected shell/area objects: {ex.Message}");
            }
        }

        internal static OperationResult<CSISapModelShellObjectDTO> GetByName<TSapModel>(
            TSapModel sapModel,
            string productName,
            string areaName,
            CSISapModelAreaGetPoints<TSapModel> getPoints,
            CSISapModelAreaGetProperty<TSapModel> getProperty,
            CSISapModelAreaGetSelected<TSapModel> getSelected)
        {
            var pointsResult = GetPoints(sapModel, productName, areaName, getPoints);
            if (!pointsResult.IsSuccess)
            {
                return OperationResult<CSISapModelShellObjectDTO>.Failure(pointsResult.Message);
            }

            var propertyResult = GetProperty(sapModel, productName, areaName, getProperty);
            if (!propertyResult.IsSuccess)
            {
                return OperationResult<CSISapModelShellObjectDTO>.Failure(propertyResult.Message);
            }

            bool selected = false;
            int selectedResult = getSelected(sapModel, areaName, ref selected);
            if (selectedResult != 0)
            {
                return OperationResult<CSISapModelShellObjectDTO>.Failure($"{productName} AreaObj.GetSelected failed for '{areaName}' (return code {selectedResult}).");
            }

            return OperationResult<CSISapModelShellObjectDTO>.Success(new CSISapModelShellObjectDTO
            {
                Name = areaName,
                PointNames = pointsResult.Data,
                PropertyName = propertyResult.Data,
                IsSelected = selected
            });
        }

        internal static OperationResult<IReadOnlyList<CSISapModelShellLoadDTO>> GetUniformLoads<TSapModel>(
            TSapModel sapModel,
            string productName,
            string areaName,
            CSISapModelAreaGetLoadUniform<TSapModel> getLoadUniform)
        {
            var validation = ShellObjectValidation.ValidateAreaName(areaName);
            if (!validation.IsSuccess)
            {
                return OperationResult<IReadOnlyList<CSISapModelShellLoadDTO>>.Failure(validation.Message);
            }

            try
            {
                int numberItems = 0;
                string[] areaNames = null;
                string[] loadPatterns = null;
                string[] coordinateSystems = null;
                int[] directions = null;
                double[] values = null;

                int result = getLoadUniform(
                    sapModel,
                    areaName,
                    ref numberItems,
                    ref areaNames,
                    ref loadPatterns,
                    ref coordinateSystems,
                    ref directions,
                    ref values);

                if (result != 0)
                {
                    return OperationResult<IReadOnlyList<CSISapModelShellLoadDTO>>.Failure($"{productName} AreaObj.GetLoadUniform failed for '{areaName}' (return code {result}).");
                }

                var loads = new List<CSISapModelShellLoadDTO>();
                for (int i = 0; i < numberItems; i++)
                {
                    loads.Add(new CSISapModelShellLoadDTO
                    {
                        AreaName = GetArrayValue(areaNames, i),
                        LoadPattern = GetArrayValue(loadPatterns, i),
                        LoadType = "Uniform",
                        CoordinateSystem = GetArrayValue(coordinateSystems, i),
                        Direction = GetArrayValue(directions, i),
                        Value = GetArrayValue(values, i)
                    });
                }

                return OperationResult<IReadOnlyList<CSISapModelShellLoadDTO>>.Success(loads);
            }
            catch (COMException ex)
            {
                return OperationResult<IReadOnlyList<CSISapModelShellLoadDTO>>.Failure($"{productName} COM error while reading shell/area uniform loads: {ex.Message}");
            }
            catch (Exception ex)
            {
                return OperationResult<IReadOnlyList<CSISapModelShellLoadDTO>>.Failure($"Failed to read shell/area uniform loads: {ex.Message}");
            }
        }

        internal static CsiWritePreview PreviewAddByPoint(IReadOnlyList<string> pointNames, string propertyName, string userName)
        {
            return Preview(
                "shells.add_by_points",
                CsiMethodRiskLevel.Low,
                false,
                true,
                $"This will add one shell/area object using {Count(pointNames)} existing point(s), property '{CleanPropertyName(propertyName)}', and name '{CleanName(userName)}'.",
                One(CleanName(userName)));
        }

        internal static OperationResult<string> AddByPoint<TSapModel>(
            TSapModel sapModel,
            string productName,
            IReadOnlyList<string> pointNames,
            string propertyName,
            string userName,
            bool confirmed,
            CsiWriteGuard writeGuard,
            CsiOperationLogger logger,
            CSISapModelReadPointCoordinates<TSapModel> getPointCoordinates,
            CSISapModelAreaAddByPoint<TSapModel> addByPoint,
            Func<TSapModel, OperationResult> refreshView)
        {
            var validation = ShellObjectValidation.ValidatePointNames(pointNames);
            if (!validation.IsSuccess)
            {
                return OperationResult<string>.Failure(validation.Message);
            }

            for (int i = 0; i < pointNames.Count; i++)
            {
                double x = 0;
                double y = 0;
                double z = 0;
                int pointResult = getPointCoordinates(sapModel, pointNames[i], ref x, ref y, ref z);
                if (pointResult != 0)
                {
                    return OperationResult<string>.Failure($"Point '{pointNames[i]}' does not exist in {productName}.");
                }
            }

            string operationName = "shells.add_by_points";
            string property = CleanPropertyName(propertyName);
            string name = CleanName(userName);
            string summary = $"points={pointNames.Count}, property={property}, userName={name}";
            IReadOnlyList<string> affected = One(name);

            var guardResult = writeGuard.ValidateWrite(operationName, CsiMethodRiskLevel.Low, confirmed, affected);
            if (!guardResult.IsSuccess)
            {
                logger.Log(productName, operationName, "Shells / Areas", "Creation", CsiMethodRiskLevel.Low, summary, affected, confirmed, false, guardResult.Message);
                return OperationResult<string>.Failure(guardResult.Message);
            }

            try
            {
                string[] points = pointNames.ToArray();
                string areaName = name;
                int result = addByPoint(sapModel, points.Length, ref points, ref areaName, property, name);
                bool success = result == 0;
                string message = success
                    ? $"Created shell/area '{areaName}'."
                    : $"{productName} AreaObj.AddByPoint failed (return code {result}).";

                if (success)
                {
                    var refresh = refreshView(sapModel);
                    if (!refresh.IsSuccess)
                    {
                        message = refresh.Message;
                        success = false;
                    }
                }

                logger.Log(productName, operationName, "Shells / Areas", "Creation", CsiMethodRiskLevel.Low, summary, One(areaName), confirmed, success, message);
                return success ? OperationResult<string>.Success(areaName, message) : OperationResult<string>.Failure(message);
            }
            catch (COMException ex)
            {
                logger.Log(productName, operationName, "Shells / Areas", "Creation", CsiMethodRiskLevel.Low, summary, affected, confirmed, false, ex.Message);
                return OperationResult<string>.Failure($"{productName} COM error while adding shell/area by points: {ex.Message}");
            }
        }

        internal static CsiWritePreview PreviewAddByCoord(IReadOnlyList<CSISapModelShellCoordinateInput> points, string propertyName, string userName)
        {
            return Preview(
                "shells.add_by_coordinates",
                CsiMethodRiskLevel.Low,
                false,
                true,
                $"This will add one shell/area object using {Count(points)} coordinate point(s), property '{CleanPropertyName(propertyName)}', and name '{CleanName(userName)}'.",
                One(CleanName(userName)));
        }

        internal static OperationResult<string> AddByCoord<TSapModel>(
            TSapModel sapModel,
            string productName,
            IReadOnlyList<CSISapModelShellCoordinateInput> points,
            string propertyName,
            string userName,
            string coordinateSystem,
            bool confirmed,
            CsiWriteGuard writeGuard,
            CsiOperationLogger logger,
            CSISapModelAreaAddByCoord<TSapModel> addByCoord,
            Func<TSapModel, OperationResult> refreshView)
        {
            var validation = ShellObjectValidation.ValidateCoordinates(points);
            if (!validation.IsSuccess)
            {
                return OperationResult<string>.Failure(validation.Message);
            }

            string operationName = "shells.add_by_coordinates";
            string property = CleanPropertyName(propertyName);
            string name = CleanName(userName);
            string cSys = string.IsNullOrWhiteSpace(coordinateSystem) ? "Global" : coordinateSystem.Trim();
            string summary = $"points={points.Count}, property={property}, userName={name}, cSys={cSys}";
            IReadOnlyList<string> affected = One(name);

            var guardResult = writeGuard.ValidateWrite(operationName, CsiMethodRiskLevel.Low, confirmed, affected);
            if (!guardResult.IsSuccess)
            {
                logger.Log(productName, operationName, "Shells / Areas", "Creation", CsiMethodRiskLevel.Low, summary, affected, confirmed, false, guardResult.Message);
                return OperationResult<string>.Failure(guardResult.Message);
            }

            try
            {
                double[] x = points.Select(p => p.X).ToArray();
                double[] y = points.Select(p => p.Y).ToArray();
                double[] z = points.Select(p => p.Z).ToArray();
                string areaName = name;
                int result = addByCoord(sapModel, points.Count, ref x, ref y, ref z, ref areaName, property, name, cSys);
                bool success = result == 0;
                string message = success
                    ? $"Created shell/area '{areaName}'."
                    : $"{productName} AreaObj.AddByCoord failed (return code {result}).";

                if (success)
                {
                    var refresh = refreshView(sapModel);
                    if (!refresh.IsSuccess)
                    {
                        message = refresh.Message;
                        success = false;
                    }
                }

                logger.Log(productName, operationName, "Shells / Areas", "Creation", CsiMethodRiskLevel.Low, summary, One(areaName), confirmed, success, message);
                return success ? OperationResult<string>.Success(areaName, message) : OperationResult<string>.Failure(message);
            }
            catch (COMException ex)
            {
                logger.Log(productName, operationName, "Shells / Areas", "Creation", CsiMethodRiskLevel.Low, summary, affected, confirmed, false, ex.Message);
                return OperationResult<string>.Failure($"{productName} COM error while adding shell/area by coordinates: {ex.Message}");
            }
        }

        internal static CsiWritePreview PreviewAssignUniformLoad(IReadOnlyList<string> areaNames, string loadPattern, double value, int direction, bool replace, string coordinateSystem)
        {
            return Preview(
                "shells.assign_uniform_load",
                CsiMethodRiskLevel.Medium,
                true,
                true,
                $"This will assign uniform load pattern '{loadPattern}' to {Count(areaNames)} shell/area object(s).",
                areaNames);
        }

        internal static OperationResult AssignUniformLoad<TSapModel>(
            TSapModel sapModel,
            string productName,
            IReadOnlyList<string> areaNames,
            string loadPattern,
            double value,
            int direction,
            bool replace,
            string coordinateSystem,
            bool confirmed,
            CsiWriteGuard writeGuard,
            CsiOperationLogger logger,
            CSISapModelAreaGetPoints<TSapModel> getPoints,
            CSISapModelAreaGetNameList<TSapModel> getLoadPatternNames,
            CSISapModelAreaSetLoadUniform<TSapModel> setLoadUniform,
            Func<TSapModel, OperationResult> refreshView)
        {
            var areaValidation = ShellObjectValidation.ValidateAreaNames(areaNames);
            if (!areaValidation.IsSuccess)
            {
                return areaValidation;
            }

            var loadValidation = ShellObjectValidation.ValidateUniformLoad(loadPattern, direction, coordinateSystem);
            if (!loadValidation.IsSuccess)
            {
                return loadValidation;
            }

            if (!NameExists(sapModel, loadPattern, getLoadPatternNames))
            {
                return OperationResult.Failure($"Load pattern '{loadPattern}' does not exist.");
            }

            for (int i = 0; i < areaNames.Count; i++)
            {
                if (!AreaExists(sapModel, areaNames[i], getPoints))
                {
                    return OperationResult.Failure($"Shell/area '{areaNames[i]}' does not exist.");
                }
            }

            string operationName = "shells.assign_uniform_load";
            string summary = $"loadPattern={loadPattern}, value={value}, direction={direction}, replace={replace}, cSys={coordinateSystem}, areas={areaNames.Count}";
            var guardResult = writeGuard.ValidateWrite(operationName, CsiMethodRiskLevel.Medium, confirmed, areaNames);
            if (!guardResult.IsSuccess)
            {
                logger.Log(productName, operationName, "Shells / Areas", "Loads", CsiMethodRiskLevel.Medium, summary, areaNames, confirmed, false, guardResult.Message);
                return guardResult;
            }

            var failed = new List<string>();
            foreach (string areaName in areaNames)
            {
                int result = setLoadUniform(sapModel, areaName, loadPattern, value, direction, replace, coordinateSystem);
                if (result != 0)
                {
                    failed.Add($"{areaName} (return code {result})");
                }
            }

            bool success = failed.Count == 0;
            string message = success
                ? $"Assigned uniform load '{loadPattern}' to {areaNames.Count} shell/area object(s)."
                : "Some shell/area uniform load assignments failed: " + string.Join(", ", failed);

            if (success)
            {
                var refresh = refreshView(sapModel);
                if (!refresh.IsSuccess)
                {
                    message = refresh.Message;
                    success = false;
                }
            }

            logger.Log(productName, operationName, "Shells / Areas", "Loads", CsiMethodRiskLevel.Medium, summary, areaNames, confirmed, success, message);
            return success ? OperationResult.Success(message) : OperationResult.Failure(message);
        }

        internal static CsiWritePreview PreviewDeleteAreas(IReadOnlyList<string> areaNames)
        {
            return Preview(
                "shells.delete",
                CsiMethodRiskLevel.High,
                true,
                true,
                $"This will delete {Count(areaNames)} shell/area object(s). This is high risk.",
                areaNames);
        }

        internal static OperationResult DeleteAreas<TSapModel>(
            TSapModel sapModel,
            string productName,
            IReadOnlyList<string> areaNames,
            bool confirmed,
            CsiWriteGuard writeGuard,
            CsiOperationLogger logger,
            CSISapModelAreaGetPoints<TSapModel> getPoints,
            CSISapModelAreaDelete<TSapModel> deleteArea,
            Func<TSapModel, OperationResult> refreshView)
        {
            var validation = ShellObjectValidation.ValidateAreaNames(areaNames);
            if (!validation.IsSuccess)
            {
                return validation;
            }

            for (int i = 0; i < areaNames.Count; i++)
            {
                if (!AreaExists(sapModel, areaNames[i], getPoints))
                {
                    return OperationResult.Failure($"Shell/area '{areaNames[i]}' does not exist.");
                }
            }

            string operationName = "shells.delete";
            string summary = $"areas={areaNames.Count}";
            var guardResult = writeGuard.ValidateWrite(operationName, CsiMethodRiskLevel.High, confirmed, areaNames);
            if (!guardResult.IsSuccess)
            {
                logger.Log(productName, operationName, "Shells / Areas", "Deletion", CsiMethodRiskLevel.High, summary, areaNames, confirmed, false, guardResult.Message);
                return guardResult;
            }

            var failed = new List<string>();
            foreach (string areaName in areaNames)
            {
                int result = deleteArea(sapModel, areaName);
                if (result != 0)
                {
                    failed.Add($"{areaName} (return code {result})");
                }
            }

            bool success = failed.Count == 0;
            string message = success
                ? $"Deleted {areaNames.Count} shell/area object(s)."
                : "Some shell/area deletes failed: " + string.Join(", ", failed);

            if (success)
            {
                var refresh = refreshView(sapModel);
                if (!refresh.IsSuccess)
                {
                    message = refresh.Message;
                    success = false;
                }
            }

            logger.Log(productName, operationName, "Shells / Areas", "Deletion", CsiMethodRiskLevel.High, summary, areaNames, confirmed, success, message);
            return success ? OperationResult.Success(message) : OperationResult.Failure(message);
        }

        internal static OperationResult CreateShellAreasFromSelectedFrames<TSapModel>(
            TSapModel sapModel,
            string productName,
            string propertyName,
            ShellCreationTolerances tolerances,
            Func<TSapModel, int> setShellCreationUnits,
            CSISapModelGetSelectedObjects<TSapModel> getSelectedObjects,
            CSISapModelGetFramePoints<TSapModel> getFramePoints,
            CSISapModelGetPointCoordinates<TSapModel> getPointCoordinates,
            CSISapModelAddAreaByCoordinates<TSapModel> addAreaByCoordinates,
            Func<TSapModel, OperationResult> refreshView)
        {
            tolerances = tolerances ?? new ShellCreationTolerances();
            propertyName = string.IsNullOrWhiteSpace(propertyName) ? "Default" : propertyName.Trim();

            try
            {
                int unitRet = setShellCreationUnits(sapModel);
                if (unitRet != 0)
                {
                    return OperationResult.Failure($"Failed to set {productName} present units for shell area creation.");
                }

                var framesResult = ReadSelectedFrameGeometries(
                    sapModel,
                    productName,
                    getSelectedObjects,
                    getFramePoints,
                    getPointCoordinates);

                if (!framesResult.IsSuccess)
                {
                    return OperationResult.Failure(framesResult.Message);
                }

                var frameGeometries = framesResult.Data;
                if (frameGeometries == null || frameGeometries.Count == 0)
                {
                    return OperationResult.Failure($"No frame objects are currently selected in {productName}.");
                }

                var faceBuildResult = ShellFaceBuilder.BuildCandidateFaces(frameGeometries, tolerances);
                if (faceBuildResult.EnrichedRealEdgeCount == 0)
                {
                    return OperationResult.Failure("No valid enriched frame graph was found from the current selection.");
                }

                if (faceBuildResult.FaceLoops == null || faceBuildResult.FaceLoops.Count == 0)
                {
                    return OperationResult.Failure("No closed shell faces were extracted from the selected frames.");
                }

                var acceptedFaces = new List<IReadOnlyList<string>>();
                var createdCount = 0;
                var skippedCount = 0;
                var shellFaceCandidates = BuildShellFaceCandidates(
                    faceBuildResult.FaceLoops,
                    faceBuildResult.PointCoordinates,
                    tolerances,
                    ref skippedCount);

                var progress = BatchProgressWindow.RunWithProgress(
                    shellFaceCandidates.Count,
                    "Creating Shell Areas From Selected Frames...",
                    ctx =>
                    {
                        foreach (var candidate in shellFaceCandidates)
                        {
                            if (ctx.IsCancellationRequested)
                            {
                                break;
                            }

                            string rejectReason;
                            if (!ShellFaceBuilder.ValidateFaceLoop(
                                    candidate.OrderedLoop,
                                    faceBuildResult.PointCoordinates,
                                    acceptedFaces,
                                    tolerances,
                                    out rejectReason))
                            {
                                skippedCount++;
                                ctx.IncrementSkipped();
                                continue;
                            }

                            var createdForLoop = CreateAreaFromLoop(
                                sapModel,
                                candidate.OrderedLoop,
                                faceBuildResult.PointCoordinates,
                                propertyName,
                                acceptedFaces,
                                tolerances,
                                addAreaByCoordinates);

                            if (createdForLoop > 0)
                            {
                                createdCount += createdForLoop;
                                ctx.IncrementRan();
                            }
                            else
                            {
                                skippedCount++;
                                ctx.IncrementSkipped();
                            }
                        }
                    });

                var refreshResult = refreshView(sapModel);
                if (!refreshResult.IsSuccess)
                {
                    return refreshResult;
                }

                var message = "Done." + Environment.NewLine +
                              $"Created: {createdCount}" + Environment.NewLine +
                              $"Skipped: {skippedCount}";

                if (progress.WasCancelled)
                {
                    message += Environment.NewLine + "Operation cancelled before all faces were processed.";
                }

                return OperationResult.Success(message);
            }
            catch (COMException ex)
            {
                return OperationResult.Failure($"{productName} COM error while creating shell areas from selected frames: {ex.Message}");
            }
            catch (Exception ex)
            {
                return OperationResult.Failure($"Failed to create shell areas from selected frames: {ex.Message}");
            }
        }

        private static OperationResult<IReadOnlyList<ShellFrameGeometry>> ReadSelectedFrameGeometries<TSapModel>(
            TSapModel sapModel,
            string productName,
            CSISapModelGetSelectedObjects<TSapModel> getSelectedObjects,
            CSISapModelGetFramePoints<TSapModel> getFramePoints,
            CSISapModelGetPointCoordinates<TSapModel> getPointCoordinates)
        {
            int numberItems = 0;
            int[] objectTypes = null;
            string[] objectNames = null;
            int selectedResult = getSelectedObjects(sapModel, ref numberItems, ref objectTypes, ref objectNames);

            if (selectedResult != 0)
            {
                return OperationResult<IReadOnlyList<ShellFrameGeometry>>.Failure($"Failed to read selected objects from {productName}.");
            }

            if (numberItems <= 0 || objectTypes == null || objectNames == null)
            {
                return OperationResult<IReadOnlyList<ShellFrameGeometry>>.Failure($"No frame objects are currently selected in {productName}.");
            }

            var frames = new List<ShellFrameGeometry>();
            var seenFrames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            for (int i = 0; i < numberItems; i++)
            {
                if (i >= objectTypes.Length || i >= objectNames.Length)
                {
                    continue;
                }

                var frameName = objectNames[i];
                if (objectTypes[i] != CSISapModelObjectTypeIds.Frame ||
                    string.IsNullOrWhiteSpace(frameName) ||
                    seenFrames.Contains(frameName))
                {
                    continue;
                }

                string p1 = string.Empty;
                string p2 = string.Empty;
                int framePointsResult = getFramePoints(sapModel, frameName, ref p1, ref p2);
                if (framePointsResult != 0 || string.IsNullOrWhiteSpace(p1) || string.IsNullOrWhiteSpace(p2))
                {
                    continue;
                }

                double x1 = 0;
                double y1 = 0;
                double z1 = 0;
                double x2 = 0;
                double y2 = 0;
                double z2 = 0;

                int p1Result = getPointCoordinates(sapModel, p1, ref x1, ref y1, ref z1);
                int p2Result = getPointCoordinates(sapModel, p2, ref x2, ref y2, ref z2);
                if (p1Result != 0 || p2Result != 0)
                {
                    continue;
                }

                seenFrames.Add(frameName);
                frames.Add(new ShellFrameGeometry
                {
                    FrameName = frameName,
                    StartPointName = p1,
                    EndPointName = p2,
                    StartPoint = new ShellGeometryPoint3D(x1, y1, z1),
                    EndPoint = new ShellGeometryPoint3D(x2, y2, z2)
                });
            }

            if (frames.Count == 0)
            {
                return OperationResult<IReadOnlyList<ShellFrameGeometry>>.Failure($"No valid frame geometry could be read from the current {productName} selection.");
            }

            return OperationResult<IReadOnlyList<ShellFrameGeometry>>.Success(frames);
        }

        private static List<ShellFaceCandidate> BuildShellFaceCandidates(
            IReadOnlyList<string[]> rawFaceLoops,
            IReadOnlyDictionary<string, ShellGeometryPoint3D> pointCoords,
            ShellCreationTolerances tolerances,
            ref int skippedCount)
        {
            var candidates = new List<ShellFaceCandidate>();
            var emptyAcceptedFaces = new List<IReadOnlyList<string>>();

            foreach (var rawLoop in rawFaceLoops)
            {
                var cleanLoop = ShellFaceBuilder.CleanLoopBoundaryXY(rawLoop, pointCoords, tolerances);
                if (cleanLoop == null || cleanLoop.Length < 3)
                {
                    skippedCount++;
                    continue;
                }

                var orderedLoop = ShellFaceBuilder.OrderLoopUpward(cleanLoop, pointCoords);
                string rejectReason;
                if (!ShellFaceBuilder.ValidateFaceLoop(
                        orderedLoop,
                        pointCoords,
                        emptyAcceptedFaces,
                        tolerances,
                        out rejectReason))
                {
                    skippedCount++;
                    continue;
                }

                candidates.Add(new ShellFaceCandidate
                {
                    OrderedLoop = orderedLoop,
                    Area = Math.Abs(ShellFaceBuilder.GetPolygonAreaXY(orderedLoop, pointCoords))
                });
            }

            var sortedCandidates = candidates
                .OrderBy(candidate => GetShellLoopPriority(candidate.OrderedLoop.Length))
                .ThenBy(candidate => candidate.OrderedLoop.Length)
                .ThenBy(candidate => candidate.Area)
                .ToList();
            return sortedCandidates;
        }

        private static int GetShellLoopPriority(int nodeCount)
        {
            if (nodeCount == 4) return 0;
            if (nodeCount == 3) return 1;
            return 2;
        }

        private static int CreateAreaFromLoop<TSapModel>(
            TSapModel sapModel,
            IReadOnlyList<string> loopPts,
            IReadOnlyDictionary<string, ShellGeometryPoint3D> pointCoords,
            string propName,
            List<IReadOnlyList<string>> acceptedFaces,
            ShellCreationTolerances tolerances,
            CSISapModelAddAreaByCoordinates<TSapModel> addAreaByCoordinates)
        {
            if (loopPts == null)
            {
                return 0;
            }

            if (loopPts.Count == 3)
            {
                if (AddAreaByNodeCoordinates(sapModel, loopPts, pointCoords, propName, addAreaByCoordinates))
                {
                    acceptedFaces.Add(loopPts.ToArray());
                    return 1;
                }

                return 0;
            }

            if (loopPts.Count == 4)
            {
                if (AddAreaByNodeCoordinates(sapModel, loopPts, pointCoords, propName, addAreaByCoordinates))
                {
                    acceptedFaces.Add(loopPts.ToArray());
                    return 1;
                }

                return SplitQuadAndCreateTwoTriangles(
                    sapModel,
                    loopPts,
                    pointCoords,
                    propName,
                    acceptedFaces,
                    tolerances,
                    addAreaByCoordinates);
            }

            return 0;
        }

        private static bool AddAreaByNodeCoordinates<TSapModel>(
            TSapModel sapModel,
            IReadOnlyList<string> nodeIds,
            IReadOnlyDictionary<string, ShellGeometryPoint3D> pointCoords,
            string propName,
            CSISapModelAddAreaByCoordinates<TSapModel> addAreaByCoordinates)
        {
            if (nodeIds == null || (nodeIds.Count != 3 && nodeIds.Count != 4))
            {
                return false;
            }

            var x = new double[nodeIds.Count];
            var y = new double[nodeIds.Count];
            var z = new double[nodeIds.Count];

            for (int i = 0; i < nodeIds.Count; i++)
            {
                ShellGeometryPoint3D point;
                if (!pointCoords.TryGetValue(nodeIds[i], out point))
                {
                    return false;
                }

                x[i] = point.X;
                y[i] = point.Y;
                z[i] = point.Z;
            }

            string areaName = string.Empty;
            int addResult = addAreaByCoordinates(sapModel, nodeIds.Count, ref x, ref y, ref z, ref areaName, propName);
            return addResult == 0;
        }

        private static int SplitQuadAndCreateTwoTriangles<TSapModel>(
            TSapModel sapModel,
            IReadOnlyList<string> quadPts,
            IReadOnlyDictionary<string, ShellGeometryPoint3D> pointCoords,
            string propName,
            List<IReadOnlyList<string>> acceptedFaces,
            ShellCreationTolerances tolerances,
            CSISapModelAddAreaByCoordinates<TSapModel> addAreaByCoordinates)
        {
            var acLen = ShellFaceBuilder.Distance3D(pointCoords[quadPts[0]], pointCoords[quadPts[2]]);
            var bdLen = ShellFaceBuilder.Distance3D(pointCoords[quadPts[1]], pointCoords[quadPts[3]]);

            string[] tri1;
            string[] tri2;

            if (acLen <= bdLen)
            {
                tri1 = new[] { quadPts[0], quadPts[1], quadPts[2] };
                tri2 = new[] { quadPts[0], quadPts[2], quadPts[3] };
            }
            else
            {
                tri1 = new[] { quadPts[0], quadPts[1], quadPts[3] };
                tri2 = new[] { quadPts[1], quadPts[2], quadPts[3] };
            }

            var tri1Up = ShellFaceBuilder.OrderLoopUpward(tri1, pointCoords);
            var tri2Up = ShellFaceBuilder.OrderLoopUpward(tri2, pointCoords);

            if (ShellFaceBuilder.IsDegenerateTriangle(tri1Up, pointCoords) ||
                ShellFaceBuilder.IsDegenerateTriangle(tri2Up, pointCoords))
            {
                return 0;
            }

            if (ShellFaceBuilder.OverlapsAcceptedFacesXY(tri1Up, acceptedFaces, pointCoords, tolerances.PointTolerance) ||
                ShellFaceBuilder.OverlapsAcceptedFacesXY(tri2Up, acceptedFaces, pointCoords, tolerances.PointTolerance))
            {
                return 0;
            }

            if (!AddAreaByNodeCoordinates(sapModel, tri1Up, pointCoords, propName, addAreaByCoordinates))
            {
                return 0;
            }

            acceptedFaces.Add(tri1Up);

            if (!AddAreaByNodeCoordinates(sapModel, tri2Up, pointCoords, propName, addAreaByCoordinates))
            {
                return 1;
            }

            acceptedFaces.Add(tri2Up);
            return 2;
        }

        private class ShellFaceCandidate
        {
            public string[] OrderedLoop { get; set; }
            public double Area { get; set; }
        }

        private static CsiWritePreview Preview(
            string operationName,
            CsiMethodRiskLevel riskLevel,
            bool requiresConfirmation,
            bool supportsDryRun,
            string summary,
            IReadOnlyList<string> affectedObjects)
        {
            return new CsiWritePreview
            {
                OperationName = operationName,
                RiskLevel = riskLevel,
                RequiresConfirmation = requiresConfirmation,
                SupportsDryRun = supportsDryRun,
                Summary = summary,
                AffectedObjects = affectedObjects ?? new string[0]
            };
        }

        private static bool AreaExists<TSapModel>(
            TSapModel sapModel,
            string areaName,
            CSISapModelAreaGetPoints<TSapModel> getPoints)
        {
            int numberPoints = 0;
            string[] pointNames = null;
            return getPoints(sapModel, areaName, ref numberPoints, ref pointNames) == 0;
        }

        private static bool NameExists<TSapModel>(
            TSapModel sapModel,
            string name,
            CSISapModelAreaGetNameList<TSapModel> getNameList)
        {
            int numberNames = 0;
            string[] names = null;
            int result = getNameList(sapModel, ref numberNames, ref names);
            return result == 0 &&
                   names != null &&
                   names.Take(numberNames).Any(item => string.Equals(item, name, StringComparison.OrdinalIgnoreCase));
        }

        private static string CleanPropertyName(string propertyName)
        {
            return string.IsNullOrWhiteSpace(propertyName) ? "Default" : propertyName.Trim();
        }

        private static string CleanName(string value)
        {
            return string.IsNullOrWhiteSpace(value) ? string.Empty : value.Trim();
        }

        private static IReadOnlyList<string> One(string value)
        {
            return string.IsNullOrWhiteSpace(value) ? new string[0] : new[] { value };
        }

        private static int Count<T>(IReadOnlyList<T> values)
        {
            return values == null ? 0 : values.Count;
        }

        private static string GetArrayValue(string[] values, int index)
        {
            return values == null || index < 0 || index >= values.Length ? string.Empty : values[index];
        }

        private static int GetArrayValue(int[] values, int index)
        {
            return values == null || index < 0 || index >= values.Length ? 0 : values[index];
        }

        private static double GetArrayValue(double[] values, int index)
        {
            return values == null || index < 0 || index >= values.Length ? 0.0 : values[index];
        }
    }
}

