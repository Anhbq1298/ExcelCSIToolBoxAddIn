using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using ExcelCSIToolBoxAddIn.Common.Results;
using ExcelCSIToolBoxAddIn.Core.Geometry;
using ExcelCSIToolBoxAddIn.Data;
using ExcelCSIToolBoxAddIn.Data.DTOs;
using ExcelCSIToolBoxAddIn.Data.Models;
using ExcelCSIToolBoxAddIn.UI.Views;

namespace ExcelCSIToolBoxAddIn.Infrastructure.CSISapModel
{
    internal delegate int CSISapModelGetSelectedObjects<TSapModel>(
        TSapModel sapModel,
        ref int numberItems,
        ref int[] objectTypes,
        ref string[] objectNames);

    internal delegate int CSISapModelGetFramePoints<TSapModel>(
        TSapModel sapModel,
        string frameName,
        ref string point1Name,
        ref string point2Name);

    internal delegate int CSISapModelGetPointCoordinates<TSapModel>(
        TSapModel sapModel,
        string pointName,
        ref double x,
        ref double y,
        ref double z);

    internal delegate int CSISapModelAddAreaByCoordinates<TSapModel>(
        TSapModel sapModel,
        int nodeCount,
        ref double[] x,
        ref double[] y,
        ref double[] z,
        ref string areaName,
        string propertyName);

    internal static class CSISapModelShellObjectService
    {
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
                    StartPoint = new ShellPoint3D(x1, y1, z1),
                    EndPoint = new ShellPoint3D(x2, y2, z2)
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
            IReadOnlyDictionary<string, ShellPoint3D> pointCoords,
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
            IReadOnlyDictionary<string, ShellPoint3D> pointCoords,
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
            IReadOnlyDictionary<string, ShellPoint3D> pointCoords,
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
                ShellPoint3D point;
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
            IReadOnlyDictionary<string, ShellPoint3D> pointCoords,
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
    }
}
