using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelCSIToolBoxAddIn.Core.Geometry
{
    public class ShellCreationTolerances
    {
        private const double KnMCToleranceScale = 1000.0;

        public double PointTolerance { get; set; } = 0.0005 * KnMCToleranceScale;
        public double CollinearTolerance { get; set; } = 0.0005 * KnMCToleranceScale;
        public double AreaTolerance { get; set; } = 0.000001 * KnMCToleranceScale * KnMCToleranceScale;
        public double IntersectTolerance { get; set; } = 0.000001 * KnMCToleranceScale;
        public double ZTolerance { get; set; } = 0.001 * KnMCToleranceScale;
    }

    public class ShellPoint3D
    {
        public ShellPoint3D(double x, double y, double z)
        {
            X = x;
            Y = y;
            Z = z;
        }

        public double X { get; }
        public double Y { get; }
        public double Z { get; }
    }

    public class ShellFrameGeometry
    {
        public string FrameName { get; set; }
        public string StartPointName { get; set; }
        public string EndPointName { get; set; }
        public ShellPoint3D StartPoint { get; set; }
        public ShellPoint3D EndPoint { get; set; }
    }

    public class ShellFaceBuildResult
    {
        public Dictionary<string, ShellPoint3D> PointCoordinates { get; set; }
        public Dictionary<string, string> NodeModelPoints { get; set; }
        public List<string[]> FaceLoops { get; set; }
        public int InitialRealEdgeCount { get; set; }
        public int EnrichedRealEdgeCount { get; set; }
        public int VirtualEdgeCount { get; set; }
        public int ExtractedFaceCount { get; set; }
        public int OuterFaceRemovedCount { get; set; }
    }

    public static class ShellFaceBuilder
    {
        private const double SplitParamTolerance = 0.000000001;
        private const double Tiny = 0.000000000001;

        public static ShellFaceBuildResult BuildCandidateFaces(
            IReadOnlyList<ShellFrameGeometry> frames,
            ShellCreationTolerances tolerances)
        {
            if (frames == null)
            {
                throw new ArgumentNullException(nameof(frames));
            }

            tolerances = tolerances ?? new ShellCreationTolerances();

            var frameMap = frames
                .Where(IsValidFrame)
                .GroupBy(f => f.FrameName, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(g => g.Key, g => g.First(), StringComparer.OrdinalIgnoreCase);

            var pointCoords = new Dictionary<string, ShellPoint3D>(StringComparer.OrdinalIgnoreCase);
            var nodeModelPoint = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var initialEdges = new Dictionary<string, Edge>(StringComparer.OrdinalIgnoreCase);
            var initialAdj = NewAdjacency();

            foreach (var frame in frameMap.Values)
            {
                AddPointIfMissing(pointCoords, frame.StartPointName, frame.StartPoint);
                AddPointIfMissing(pointCoords, frame.EndPointName, frame.EndPoint);
                if (!nodeModelPoint.ContainsKey(frame.StartPointName)) nodeModelPoint.Add(frame.StartPointName, frame.StartPointName);
                if (!nodeModelPoint.ContainsKey(frame.EndPointName)) nodeModelPoint.Add(frame.EndPointName, frame.EndPointName);
                AddUndirectedEdge(initialEdges, initialAdj, frame.StartPointName, frame.EndPointName);
            }

            var frameSplitMap = InitializeFrameSplitMap(frameMap);
            EnrichRealGraphByGeometricIntersections(
                frameMap,
                frameSplitMap,
                pointCoords,
                nodeModelPoint,
                tolerances.PointTolerance,
                tolerances.IntersectTolerance,
                tolerances.ZTolerance);

            var enrichedEdges = new Dictionary<string, Edge>(StringComparer.OrdinalIgnoreCase);
            var enrichedAdj = NewAdjacency();
            foreach (var frameName in frameSplitMap.Keys)
            {
                AddSplitEdgesForFrame(frameName, frameSplitMap, enrichedEdges, enrichedAdj);
            }

            var virtualEdges = new Dictionary<string, Edge>(StringComparer.OrdinalIgnoreCase);
            var virtualAdj = NewAdjacency();
            BuildVirtualGraph(enrichedAdj, pointCoords, virtualEdges, virtualAdj, tolerances.CollinearTolerance);

            var sortedNeighbors = BuildSortedNeighbors(virtualAdj, pointCoords);
            var faces = ExtractFacesFromPlanarGraph(
                virtualEdges,
                virtualAdj,
                sortedNeighbors,
                pointCoords,
                tolerances.AreaTolerance);

            var extractedFaceCount = faces.Count;
            var outerRemovedCount = RemoveOuterFace(faces);

            return new ShellFaceBuildResult
            {
                PointCoordinates = pointCoords,
                NodeModelPoints = nodeModelPoint,
                FaceLoops = faces
                    .OrderBy(f => Math.Abs(f.Area))
                    .Select(f => f.Loop)
                    .ToList(),
                InitialRealEdgeCount = initialEdges.Count,
                EnrichedRealEdgeCount = enrichedEdges.Count,
                VirtualEdgeCount = virtualEdges.Count,
                ExtractedFaceCount = extractedFaceCount,
                OuterFaceRemovedCount = outerRemovedCount
            };
        }

        public static string[] CleanLoopBoundaryXY(
            IReadOnlyList<string> loopPts,
            IReadOnlyDictionary<string, ShellPoint3D> pointCoords,
            ShellCreationTolerances tolerances)
        {
            if (loopPts == null || loopPts.Count == 0)
            {
                return null;
            }

            tolerances = tolerances ?? new ShellCreationTolerances();

            var cleaned = RemoveConsecutiveDuplicateVertices(loopPts);
            if (cleaned == null || cleaned.Length == 0)
            {
                return null;
            }

            cleaned = RemoveClosingDuplicateVertex(cleaned, pointCoords, tolerances.PointTolerance);
            if (cleaned == null || cleaned.Length == 0)
            {
                return null;
            }

            cleaned = FixCollinearSplitPointsOnBoundary(cleaned, pointCoords, tolerances.CollinearTolerance);
            return cleaned != null && cleaned.Length >= 3 ? cleaned : null;
        }

        public static string[] OrderLoopUpward(
            IReadOnlyList<string> loopPts,
            IReadOnlyDictionary<string, ShellPoint3D> pointCoords)
        {
            if (loopPts == null || loopPts.Count == 0)
            {
                return null;
            }

            var arr = loopPts.Select(p => p).ToArray();
            var nz = 0.0;

            for (var i = 0; i < arr.Length; i++)
            {
                var pCur = pointCoords[arr[i]];
                var pNext = pointCoords[arr[(i + 1) % arr.Length]];
                nz += (pCur.X - pNext.X) * (pCur.Y + pNext.Y);
            }

            if (nz >= 0.0)
            {
                return arr;
            }

            Array.Reverse(arr);
            return arr;
        }

        public static bool ValidateFaceLoop(
            IReadOnlyList<string> loopPts,
            IReadOnlyDictionary<string, ShellPoint3D> pointCoords,
            IReadOnlyList<IReadOnlyList<string>> acceptedFaces,
            ShellCreationTolerances tolerances,
            out string rejectReason)
        {
            tolerances = tolerances ?? new ShellCreationTolerances();
            rejectReason = string.Empty;

            if (loopPts == null || loopPts.Count < 3)
            {
                rejectReason = "Too few vertices";
                return false;
            }

            if (Math.Abs(PolygonAreaXY(loopPts, pointCoords)) <= tolerances.AreaTolerance)
            {
                rejectReason = "Zero area";
                return false;
            }

            if (OverlapsAcceptedFacesXY(loopPts, acceptedFaces, pointCoords, tolerances.PointTolerance))
            {
                rejectReason = "Overlap skipped";
                return false;
            }

            return true;
        }

        public static bool OverlapsAcceptedFacesXY(
            IReadOnlyList<string> candidateLoop,
            IReadOnlyList<IReadOnlyList<string>> acceptedFaces,
            IReadOnlyDictionary<string, ShellPoint3D> pointCoords,
            double pointTolerance)
        {
            if (acceptedFaces == null)
            {
                return false;
            }

            foreach (var existingLoop in acceptedFaces)
            {
                if (PolygonsOverlapXY(candidateLoop, existingLoop, pointCoords, pointTolerance))
                {
                    return true;
                }
            }

            return false;
        }

        public static bool IsDegenerateTriangle(
            IReadOnlyList<string> triangle,
            IReadOnlyDictionary<string, ShellPoint3D> pointCoords)
        {
            if (triangle == null || triangle.Count != 3)
            {
                return true;
            }

            var p1 = pointCoords[triangle[0]];
            var p2 = pointCoords[triangle[1]];
            var p3 = pointCoords[triangle[2]];

            var ux = p2.X - p1.X;
            var uy = p2.Y - p1.Y;
            var uz = p2.Z - p1.Z;
            var vx = p3.X - p1.X;
            var vy = p3.Y - p1.Y;
            var vz = p3.Z - p1.Z;

            var nx = uy * vz - uz * vy;
            var ny = uz * vx - ux * vz;
            var nz = ux * vy - uy * vx;
            var normalLength = Math.Sqrt(nx * nx + ny * ny + nz * nz);
            return normalLength <= 0.000000001;
        }

        public static double Distance3D(ShellPoint3D p1, ShellPoint3D p2)
        {
            return Distance3DXYZ(p1.X, p1.Y, p1.Z, p2.X, p2.Y, p2.Z);
        }

        private static bool IsValidFrame(ShellFrameGeometry frame)
        {
            return frame != null &&
                   !string.IsNullOrWhiteSpace(frame.FrameName) &&
                   !string.IsNullOrWhiteSpace(frame.StartPointName) &&
                   !string.IsNullOrWhiteSpace(frame.EndPointName) &&
                   frame.StartPoint != null &&
                   frame.EndPoint != null &&
                   !string.Equals(frame.StartPointName, frame.EndPointName, StringComparison.OrdinalIgnoreCase);
        }

        private static Dictionary<string, HashSet<string>> NewAdjacency()
        {
            return new Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase);
        }

        private static void AddPointIfMissing(Dictionary<string, ShellPoint3D> pointCoords, string pointName, ShellPoint3D point)
        {
            if (!pointCoords.ContainsKey(pointName))
            {
                pointCoords.Add(pointName, point);
            }
        }

        private static Dictionary<string, List<SplitParam>> InitializeFrameSplitMap(
            Dictionary<string, ShellFrameGeometry> frameMap)
        {
            var frameSplitMap = new Dictionary<string, List<SplitParam>>(StringComparer.OrdinalIgnoreCase);

            foreach (var frame in frameMap.Values)
            {
                AddFrameSplitParam(frameSplitMap, frame.FrameName, 0.0, frame.StartPointName);
                AddFrameSplitParam(frameSplitMap, frame.FrameName, 1.0, frame.EndPointName);
            }

            return frameSplitMap;
        }

        private static void EnrichRealGraphByGeometricIntersections(
            Dictionary<string, ShellFrameGeometry> frameMap,
            Dictionary<string, List<SplitParam>> frameSplitMap,
            Dictionary<string, ShellPoint3D> pointCoords,
            Dictionary<string, string> nodeModelPoint,
            double pointTol,
            double intersectTol,
            double zTol)
        {
            var frames = frameMap.Values.ToList();

            for (var i = 0; i < frames.Count - 1; i++)
            {
                var f1 = frames[i];

                for (var j = i + 1; j < frames.Count; j++)
                {
                    var f2 = frames[j];
                    AddEndpointProjectionSplits(f1, f2, frameSplitMap, pointTol, zTol);

                    double ix;
                    double iy;
                    double t1;
                    double t2;

                    if (!SegmentIntersectionXYWithParams(
                            f1.StartPoint.X, f1.StartPoint.Y, f1.EndPoint.X, f1.EndPoint.Y,
                            f2.StartPoint.X, f2.StartPoint.Y, f2.EndPoint.X, f2.EndPoint.Y,
                            intersectTol,
                            out ix,
                            out iy,
                            out t1,
                            out t2))
                    {
                        continue;
                    }

                    if (FramesShareBothEndpoints(
                            f1.StartPointName,
                            f1.EndPointName,
                            f2.StartPointName,
                            f2.EndPointName))
                    {
                        continue;
                    }

                    var zOnF1 = f1.StartPoint.Z + t1 * (f1.EndPoint.Z - f1.StartPoint.Z);
                    var zOnF2 = f2.StartPoint.Z + t2 * (f2.EndPoint.Z - f2.StartPoint.Z);
                    if (Math.Abs(zOnF1 - zOnF2) > zTol)
                    {
                        continue;
                    }

                    var resolvedNodeId = ResolveExistingNodeAtIntersection(
                        f1.StartPointName,
                        f1.EndPointName,
                        f2.StartPointName,
                        f2.EndPointName,
                        t1,
                        t2,
                        pointTol);

                    var nodeId = string.IsNullOrWhiteSpace(resolvedNodeId)
                        ? GetOrCreateGeomIntersectionNode(pointCoords, nodeModelPoint, ix, iy, 0.5 * (zOnF1 + zOnF2), pointTol, zTol)
                        : resolvedNodeId;

                    AddFrameSplitParam(frameSplitMap, f1.FrameName, Clamp01(t1), nodeId);
                    AddFrameSplitParam(frameSplitMap, f2.FrameName, Clamp01(t2), nodeId);
                }
            }
        }

        private static void AddEndpointProjectionSplits(
            ShellFrameGeometry f1,
            ShellFrameGeometry f2,
            Dictionary<string, List<SplitParam>> frameSplitMap,
            double pointTol,
            double zTol)
        {
            AddEndpointToFrameIfOnSegment(f1.StartPointName, f1.StartPoint, f2, frameSplitMap, pointTol, zTol);
            AddEndpointToFrameIfOnSegment(f1.EndPointName, f1.EndPoint, f2, frameSplitMap, pointTol, zTol);
            AddEndpointToFrameIfOnSegment(f2.StartPointName, f2.StartPoint, f1, frameSplitMap, pointTol, zTol);
            AddEndpointToFrameIfOnSegment(f2.EndPointName, f2.EndPoint, f1, frameSplitMap, pointTol, zTol);
        }

        private static void AddEndpointToFrameIfOnSegment(
            string endpointName,
            ShellPoint3D endpoint,
            ShellFrameGeometry targetFrame,
            Dictionary<string, List<SplitParam>> frameSplitMap,
            double pointTol,
            double zTol)
        {
            if (string.Equals(endpointName, targetFrame.StartPointName, StringComparison.OrdinalIgnoreCase) ||
                string.Equals(endpointName, targetFrame.EndPointName, StringComparison.OrdinalIgnoreCase))
            {
                return;
            }

            double t;
            if (!TryGetPointProjectionParameterOnFrameXY(endpoint, targetFrame, pointTol, out t))
            {
                return;
            }

            var zOnTarget = targetFrame.StartPoint.Z + t * (targetFrame.EndPoint.Z - targetFrame.StartPoint.Z);
            if (Math.Abs(endpoint.Z - zOnTarget) > zTol)
            {
                return;
            }

            AddFrameSplitParam(frameSplitMap, targetFrame.FrameName, Clamp01(t), endpointName);
        }

        private static bool TryGetPointProjectionParameterOnFrameXY(
            ShellPoint3D point,
            ShellFrameGeometry frame,
            double pointTol,
            out double t)
        {
            t = 0.0;

            var dx = frame.EndPoint.X - frame.StartPoint.X;
            var dy = frame.EndPoint.Y - frame.StartPoint.Y;
            var lenSq = dx * dx + dy * dy;
            if (lenSq <= 0.0)
            {
                return false;
            }

            t = ((point.X - frame.StartPoint.X) * dx + (point.Y - frame.StartPoint.Y) * dy) / lenSq;
            var len = Math.Sqrt(lenSq);
            var paramTol = pointTol / len;

            if (t < -paramTol || t > 1.0 + paramTol)
            {
                return false;
            }

            var projectedX = frame.StartPoint.X + t * dx;
            var projectedY = frame.StartPoint.Y + t * dy;
            return Distance2D(point.X, point.Y, projectedX, projectedY) <= pointTol;
        }

        private static bool FramesShareBothEndpoints(string p1, string p2, string q1, string q2)
        {
            return (string.Equals(p1, q1, StringComparison.OrdinalIgnoreCase) &&
                    string.Equals(p2, q2, StringComparison.OrdinalIgnoreCase)) ||
                   (string.Equals(p1, q2, StringComparison.OrdinalIgnoreCase) &&
                    string.Equals(p2, q1, StringComparison.OrdinalIgnoreCase));
        }

        private static void AddFrameSplitParam(
            Dictionary<string, List<SplitParam>> frameSplitMap,
            string frameName,
            double t,
            string nodeId)
        {
            List<SplitParam> splitParams;
            if (!frameSplitMap.TryGetValue(frameName, out splitParams))
            {
                splitParams = new List<SplitParam>();
                frameSplitMap.Add(frameName, splitParams);
            }

            if (splitParams.Any(item => Math.Abs(item.T - t) <= SplitParamTolerance))
            {
                return;
            }

            splitParams.Add(new SplitParam { T = t, NodeId = nodeId });
        }

        private static void AddSplitEdgesForFrame(
            string frameName,
            Dictionary<string, List<SplitParam>> frameSplitMap,
            Dictionary<string, Edge> edgeDict,
            Dictionary<string, HashSet<string>> adjacency)
        {
            List<SplitParam> splitParams;
            if (!frameSplitMap.TryGetValue(frameName, out splitParams) || splitParams.Count < 2)
            {
                return;
            }

            var sorted = splitParams.OrderBy(item => item.T).ToList();

            for (var i = 0; i < sorted.Count - 1; i++)
            {
                var n1 = sorted[i].NodeId;
                var n2 = sorted[i + 1].NodeId;
                if (!string.Equals(n1, n2, StringComparison.OrdinalIgnoreCase))
                {
                    AddUndirectedEdge(edgeDict, adjacency, n1, n2);
                }
            }
        }

        private static string ResolveExistingNodeAtIntersection(
            string p1Name,
            string p2Name,
            string q1Name,
            string q2Name,
            double t1,
            double t2,
            double pointTol)
        {
            if (Math.Abs(t1) <= pointTol) return p1Name;
            if (Math.Abs(t1 - 1.0) <= pointTol) return p2Name;
            if (Math.Abs(t2) <= pointTol) return q1Name;
            if (Math.Abs(t2 - 1.0) <= pointTol) return q2Name;
            return string.Empty;
        }

        private static string GetOrCreateGeomIntersectionNode(
            Dictionary<string, ShellPoint3D> pointCoords,
            Dictionary<string, string> nodeModelPoint,
            double x,
            double y,
            double z,
            double pointTol,
            double zTol)
        {
            var tol3D = Math.Max(pointTol, zTol);

            foreach (var item in pointCoords)
            {
                var p = item.Value;
                if (Distance3DXYZ(x, y, z, p.X, p.Y, p.Z) <= tol3D)
                {
                    return item.Key;
                }
            }

            var id = "IX_" + (pointCoords.Count + 1);
            while (pointCoords.ContainsKey(id))
            {
                id = "IX_" + (pointCoords.Count + 1) + "_" + Guid.NewGuid().ToString("N").Substring(0, 6);
            }

            pointCoords.Add(id, new ShellPoint3D(x, y, z));
            nodeModelPoint.Add(id, string.Empty);
            return id;
        }

        private static void BuildVirtualGraph(
            Dictionary<string, HashSet<string>> realAdj,
            Dictionary<string, ShellPoint3D> pointCoords,
            Dictionary<string, Edge> virtualEdgeDict,
            Dictionary<string, HashSet<string>> virtualAdj,
            double collinearTol)
        {
            var chainPointDict = new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);

            foreach (var p in realAdj.Keys)
            {
                chainPointDict[p] = IsDegree2CollinearPoint(p, realAdj, pointCoords, collinearTol);
            }

            foreach (var p in realAdj.Keys)
            {
                if (chainPointDict[p])
                {
                    continue;
                }

                foreach (var n in realAdj[p])
                {
                    var endP = FollowChainUntilRetained(p, n, realAdj, pointCoords, chainPointDict, collinearTol);
                    if (!string.IsNullOrWhiteSpace(endP) &&
                        !string.Equals(p, endP, StringComparison.OrdinalIgnoreCase))
                    {
                        AddUndirectedEdge(virtualEdgeDict, virtualAdj, p, endP);
                    }
                }
            }
        }

        private static bool IsDegree2CollinearPoint(
            string p,
            Dictionary<string, HashSet<string>> adjacency,
            Dictionary<string, ShellPoint3D> pointCoords,
            double collinearTol)
        {
            HashSet<string> nbrs;
            if (!adjacency.TryGetValue(p, out nbrs) || nbrs.Count != 2)
            {
                return false;
            }

            var arr = nbrs.ToArray();
            return IsMiddlePointOnSegmentXY(arr[0], p, arr[1], pointCoords, collinearTol);
        }

        private static string FollowChainUntilRetained(
            string prevPoint,
            string currentPoint,
            Dictionary<string, HashSet<string>> adjacency,
            Dictionary<string, ShellPoint3D> pointCoords,
            Dictionary<string, bool> chainPointDict,
            double collinearTol)
        {
            var prevP = prevPoint;
            var curP = currentPoint;

            while (true)
            {
                HashSet<string> neighbors;
                if (!adjacency.TryGetValue(curP, out neighbors))
                {
                    break;
                }

                bool isChainPoint;
                if (!chainPointDict.TryGetValue(curP, out isChainPoint) || !isChainPoint)
                {
                    break;
                }

                var nextP = GetOtherNeighbor(neighbors, prevP);
                if (string.IsNullOrWhiteSpace(nextP))
                {
                    break;
                }

                if (!IsMiddlePointOnSegmentXY(prevP, curP, nextP, pointCoords, collinearTol))
                {
                    break;
                }

                prevP = curP;
                curP = nextP;
            }

            return curP;
        }

        private static Dictionary<string, string[]> BuildSortedNeighbors(
            Dictionary<string, HashSet<string>> adjacency,
            Dictionary<string, ShellPoint3D> pointCoords)
        {
            var sortedNbrs = new Dictionary<string, string[]>(StringComparer.OrdinalIgnoreCase);

            foreach (var item in adjacency)
            {
                var basePoint = pointCoords[item.Key];
                sortedNbrs[item.Key] = item.Value
                    .Select(n => new
                    {
                        NodeId = n,
                        Angle = Atan2Safe(pointCoords[n].Y - basePoint.Y, pointCoords[n].X - basePoint.X)
                    })
                    .OrderBy(n => n.Angle)
                    .Select(n => n.NodeId)
                    .ToArray();
            }

            return sortedNbrs;
        }

        private static List<FaceLoop> ExtractFacesFromPlanarGraph(
            Dictionary<string, Edge> edgeDict,
            Dictionary<string, HashSet<string>> adjacency,
            Dictionary<string, string[]> sortedNbrs,
            Dictionary<string, ShellPoint3D> pointCoords,
            double areaTol)
        {
            var faceDict = new Dictionary<string, FaceLoop>(StringComparer.OrdinalIgnoreCase);
            var visited = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (var edge in edgeDict.Values)
            {
                ExtractDirectedFace(edge.P1, edge.P2, adjacency, sortedNbrs, pointCoords, areaTol, visited, faceDict);
                ExtractDirectedFace(edge.P2, edge.P1, adjacency, sortedNbrs, pointCoords, areaTol, visited, faceDict);
            }

            return faceDict.Values.ToList();
        }

        private static void ExtractDirectedFace(
            string p1,
            string p2,
            Dictionary<string, HashSet<string>> adjacency,
            Dictionary<string, string[]> sortedNbrs,
            Dictionary<string, ShellPoint3D> pointCoords,
            double areaTol,
            HashSet<string> visited,
            Dictionary<string, FaceLoop> faceDict)
        {
            var directedKey = DirectedEdgeKey(p1, p2);
            if (visited.Contains(directedKey))
            {
                return;
            }

            var loopPts = WalkFace(p1, p2, adjacency, sortedNbrs, visited);
            if (loopPts == null || loopPts.Length < 3)
            {
                return;
            }

            var faceKey = CanonicalLoopKeyGeneric(loopPts);
            if (faceDict.ContainsKey(faceKey))
            {
                return;
            }

            var areaVal = PolygonAreaXY(loopPts, pointCoords);
            if (Math.Abs(areaVal) > areaTol)
            {
                faceDict.Add(faceKey, new FaceLoop { Key = faceKey, Loop = loopPts, Area = areaVal });
            }
        }

        private static string[] WalkFace(
            string startU,
            string startV,
            Dictionary<string, HashSet<string>> adjacency,
            Dictionary<string, string[]> sortedNbrs,
            HashSet<string> visited)
        {
            var loopList = new List<string>();
            var curU = startU;
            var curV = startV;
            var safeCounter = 0;
            var maxSteps = Math.Max(1, adjacency.Count * 10);

            do
            {
                safeCounter++;
                if (safeCounter > maxSteps)
                {
                    return null;
                }

                visited.Add(DirectedEdgeKey(curU, curV));
                loopList.Add(curU);

                var nextW = GetPrevNeighborCCW(curV, curU, sortedNbrs);
                if (string.IsNullOrWhiteSpace(nextW))
                {
                    return null;
                }

                curU = curV;
                curV = nextW;
            }
            while (!string.Equals(curU, startU, StringComparison.OrdinalIgnoreCase) ||
                   !string.Equals(curV, startV, StringComparison.OrdinalIgnoreCase));

            return loopList.ToArray();
        }

        private static string GetPrevNeighborCCW(
            string vertex,
            string incomingFrom,
            Dictionary<string, string[]> sortedNbrs)
        {
            string[] arr;
            if (!sortedNbrs.TryGetValue(vertex, out arr))
            {
                return string.Empty;
            }

            var idx = Array.FindIndex(arr, item => string.Equals(item, incomingFrom, StringComparison.OrdinalIgnoreCase));
            if (idx < 0)
            {
                return string.Empty;
            }

            return idx == 0 ? arr[arr.Length - 1] : arr[idx - 1];
        }

        private static int RemoveOuterFace(List<FaceLoop> faces)
        {
            if (faces.Count <= 1)
            {
                return 0;
            }

            var outerFace = faces.OrderByDescending(f => Math.Abs(f.Area)).First();
            faces.Remove(outerFace);
            return 1;
        }

        private static string[] RemoveConsecutiveDuplicateVertices(IReadOnlyList<string> loopPts)
        {
            var result = new List<string>();
            string previous = null;

            foreach (var point in loopPts)
            {
                if (previous == null || !string.Equals(point, previous, StringComparison.OrdinalIgnoreCase))
                {
                    result.Add(point);
                    previous = point;
                }
            }

            return result.Count > 0 ? result.ToArray() : null;
        }

        private static string[] RemoveClosingDuplicateVertex(
            string[] loopPts,
            IReadOnlyDictionary<string, ShellPoint3D> pointCoords,
            double pointTol)
        {
            if (loopPts == null || loopPts.Length == 0)
            {
                return null;
            }

            var first = pointCoords[loopPts[0]];
            var last = pointCoords[loopPts[loopPts.Length - 1]];

            if (Distance2D(first.X, first.Y, last.X, last.Y) > pointTol)
            {
                return loopPts;
            }

            if (loopPts.Length - 1 <= 0)
            {
                return null;
            }

            return loopPts.Take(loopPts.Length - 1).ToArray();
        }

        private static string[] FixCollinearSplitPointsOnBoundary(
            string[] loopPts,
            IReadOnlyDictionary<string, ShellPoint3D> pointCoords,
            double tol)
        {
            if (loopPts == null)
            {
                return null;
            }

            var current = loopPts;
            bool changed;

            do
            {
                current = RemoveOnePassCollinearSplitPoints(current, pointCoords, tol, out changed);
                if (current == null || current.Length < 3)
                {
                    return current;
                }
            }
            while (changed);

            return current;
        }

        private static string[] RemoveOnePassCollinearSplitPoints(
            string[] loopPts,
            IReadOnlyDictionary<string, ShellPoint3D> pointCoords,
            double tol,
            out bool changed)
        {
            changed = false;

            if (loopPts == null || loopPts.Length < 3)
            {
                return loopPts;
            }

            var keep = Enumerable.Repeat(true, loopPts.Length).ToArray();

            for (var i = 0; i < loopPts.Length; i++)
            {
                var prevI = (i - 1 + loopPts.Length) % loopPts.Length;
                var nextI = (i + 1) % loopPts.Length;
                if (IsMiddlePointOnSegmentXY(loopPts[prevI], loopPts[i], loopPts[nextI], pointCoords, tol))
                {
                    keep[i] = false;
                    changed = true;
                }
            }

            var result = new List<string>();
            for (var i = 0; i < loopPts.Length; i++)
            {
                if (keep[i])
                {
                    result.Add(loopPts[i]);
                }
            }

            return result.Count > 0 ? result.ToArray() : null;
        }

        private static bool PolygonsOverlapXY(
            IReadOnlyList<string> loop1,
            IReadOnlyList<string> loop2,
            IReadOnlyDictionary<string, ShellPoint3D> pointCoords,
            double pointTol)
        {
            if (PolygonsHaveProperEdgeIntersectionXY(loop1, loop2, pointCoords, pointTol))
            {
                return true;
            }

            double c1x;
            double c1y;
            GetPolygonCentroidXY(loop1, pointCoords, out c1x, out c1y);
            if (PointInsidePolygonXY(c1x, c1y, loop2, pointCoords) &&
                !PointOnPolygonBoundaryXY(c1x, c1y, loop2, pointCoords, pointTol))
            {
                return true;
            }

            double c2x;
            double c2y;
            GetPolygonCentroidXY(loop2, pointCoords, out c2x, out c2y);
            return PointInsidePolygonXY(c2x, c2y, loop1, pointCoords) &&
                   !PointOnPolygonBoundaryXY(c2x, c2y, loop1, pointCoords, pointTol);
        }

        private static bool PolygonsHaveProperEdgeIntersectionXY(
            IReadOnlyList<string> loop1,
            IReadOnlyList<string> loop2,
            IReadOnlyDictionary<string, ShellPoint3D> pointCoords,
            double pointTol)
        {
            for (var i = 0; i < loop1.Count; i++)
            {
                var a1 = pointCoords[loop1[i]];
                var a2 = pointCoords[loop1[(i + 1) % loop1.Count]];

                for (var j = 0; j < loop2.Count; j++)
                {
                    var b1 = pointCoords[loop2[j]];
                    var b2 = pointCoords[loop2[(j + 1) % loop2.Count]];

                    if (ProperSegmentsIntersectXY(a1.X, a1.Y, a2.X, a2.Y, b1.X, b1.Y, b2.X, b2.Y, pointTol))
                    {
                        return true;
                    }
                }
            }

            return false;
        }

        private static bool ProperSegmentsIntersectXY(
            double x1,
            double y1,
            double x2,
            double y2,
            double x3,
            double y3,
            double x4,
            double y4,
            double tol)
        {
            var d1 = Orient2D(x1, y1, x2, y2, x3, y3);
            var d2 = Orient2D(x1, y1, x2, y2, x4, y4);
            var d3 = Orient2D(x3, y3, x4, y4, x1, y1);
            var d4 = Orient2D(x3, y3, x4, y4, x2, y2);

            if (Math.Abs(d1) <= tol || Math.Abs(d2) <= tol || Math.Abs(d3) <= tol || Math.Abs(d4) <= tol)
            {
                return false;
            }

            return ((d1 > 0.0 && d2 < 0.0) || (d1 < 0.0 && d2 > 0.0)) &&
                   ((d3 > 0.0 && d4 < 0.0) || (d3 < 0.0 && d4 > 0.0));
        }

        private static bool PointInsidePolygonXY(
            double px,
            double py,
            IReadOnlyList<string> loopPts,
            IReadOnlyDictionary<string, ShellPoint3D> pointCoords)
        {
            var inside = false;
            var j = loopPts.Count - 1;

            for (var i = 0; i < loopPts.Count; i++)
            {
                var pi = pointCoords[loopPts[i]];
                var pj = pointCoords[loopPts[j]];

                var intersect = ((pi.Y > py) != (pj.Y > py)) &&
                                (px < (pj.X - pi.X) * (py - pi.Y) / (pj.Y - pi.Y + Tiny) + pi.X);

                if (intersect)
                {
                    inside = !inside;
                }

                j = i;
            }

            return inside;
        }

        private static bool PointOnPolygonBoundaryXY(
            double px,
            double py,
            IReadOnlyList<string> loopPts,
            IReadOnlyDictionary<string, ShellPoint3D> pointCoords,
            double tol)
        {
            for (var i = 0; i < loopPts.Count; i++)
            {
                var p1 = pointCoords[loopPts[i]];
                var p2 = pointCoords[loopPts[(i + 1) % loopPts.Count]];

                if (PointOnSegmentXY(px, py, p1.X, p1.Y, p2.X, p2.Y, tol))
                {
                    return true;
                }
            }

            return false;
        }

        private static bool PointOnSegmentXY(
            double px,
            double py,
            double x1,
            double y1,
            double x2,
            double y2,
            double tol)
        {
            var distLine = DistancePointToLineXY(px, py, x1, y1, x2, y2);
            if (distLine > tol)
            {
                return false;
            }

            var dotVal = (px - x1) * (x2 - x1) + (py - y1) * (y2 - y1);
            if (dotVal < -tol)
            {
                return false;
            }

            var lenSq = (x2 - x1) * (x2 - x1) + (y2 - y1) * (y2 - y1);
            return dotVal - lenSq <= tol;
        }

        private static void GetPolygonCentroidXY(
            IReadOnlyList<string> loopPts,
            IReadOnlyDictionary<string, ShellPoint3D> pointCoords,
            out double cx,
            out double cy)
        {
            cx = 0.0;
            cy = 0.0;

            foreach (var pointId in loopPts)
            {
                var p = pointCoords[pointId];
                cx += p.X;
                cy += p.Y;
            }

            cx /= loopPts.Count;
            cy /= loopPts.Count;
        }

        private static void AddUndirectedEdge(
            Dictionary<string, Edge> edgeDict,
            Dictionary<string, HashSet<string>> adjacency,
            string p1,
            string p2)
        {
            var eKey = EdgeKey(p1, p2);
            if (!edgeDict.ContainsKey(eKey))
            {
                edgeDict.Add(eKey, new Edge { P1 = p1, P2 = p2 });
            }

            if (!adjacency.ContainsKey(p1))
            {
                adjacency.Add(p1, new HashSet<string>(StringComparer.OrdinalIgnoreCase));
            }

            if (!adjacency.ContainsKey(p2))
            {
                adjacency.Add(p2, new HashSet<string>(StringComparer.OrdinalIgnoreCase));
            }

            adjacency[p1].Add(p2);
            adjacency[p2].Add(p1);
        }

        private static string EdgeKey(string p1, string p2)
        {
            return string.Compare(p1, p2, StringComparison.OrdinalIgnoreCase) < 0
                ? p1 + "|" + p2
                : p2 + "|" + p1;
        }

        private static string DirectedEdgeKey(string p1, string p2)
        {
            return p1 + ">>" + p2;
        }

        private static string GetOtherNeighbor(HashSet<string> neighborSet, string knownNeighbor)
        {
            return neighborSet.FirstOrDefault(k => !string.Equals(k, knownNeighbor, StringComparison.OrdinalIgnoreCase));
        }

        private static bool SegmentIntersectionXYWithParams(
            double x1,
            double y1,
            double x2,
            double y2,
            double x3,
            double y3,
            double x4,
            double y4,
            double tol,
            out double ix,
            out double iy,
            out double t1,
            out double t2)
        {
            ix = 0.0;
            iy = 0.0;
            t1 = 0.0;
            t2 = 0.0;

            var dx12 = x2 - x1;
            var dy12 = y2 - y1;
            var dx34 = x4 - x3;
            var dy34 = y4 - y3;
            var den = dx12 * dy34 - dy12 * dx34;

            if (Math.Abs(den) <= tol)
            {
                return false;
            }

            t1 = ((x3 - x1) * dy34 - (y3 - y1) * dx34) / den;
            t2 = ((x3 - x1) * dy12 - (y3 - y1) * dx12) / den;

            if (t1 < -tol || t1 > 1.0 + tol || t2 < -tol || t2 > 1.0 + tol)
            {
                return false;
            }

            t1 = Clamp01(t1);
            t2 = Clamp01(t2);
            ix = x1 + t1 * dx12;
            iy = y1 + t1 * dy12;
            return true;
        }

        private static double Clamp01(double t)
        {
            if (t < 0.0) return 0.0;
            if (t > 1.0) return 1.0;
            return t;
        }

        private static double PolygonAreaXY(
            IReadOnlyList<string> loopPts,
            IReadOnlyDictionary<string, ShellPoint3D> pointCoords)
        {
            var area2 = 0.0;

            for (var i = 0; i < loopPts.Count; i++)
            {
                var p1 = pointCoords[loopPts[i]];
                var p2 = pointCoords[loopPts[(i + 1) % loopPts.Count]];
                area2 += p1.X * p2.Y - p2.X * p1.Y;
            }

            return 0.5 * area2;
        }

        private static double Distance3DXYZ(double x1, double y1, double z1, double x2, double y2, double z2)
        {
            return Math.Sqrt((x2 - x1) * (x2 - x1) + (y2 - y1) * (y2 - y1) + (z2 - z1) * (z2 - z1));
        }

        private static double Distance2D(double x1, double y1, double x2, double y2)
        {
            return Math.Sqrt((x2 - x1) * (x2 - x1) + (y2 - y1) * (y2 - y1));
        }

        private static double DistancePointToLineXY(double px, double py, double x1, double y1, double x2, double y2)
        {
            var dx = x2 - x1;
            var dy = y2 - y1;
            var lenSq = dx * dx + dy * dy;

            if (lenSq <= 0.0)
            {
                return Math.Sqrt((px - x1) * (px - x1) + (py - y1) * (py - y1));
            }

            var crossVal = Math.Abs((px - x1) * dy - (py - y1) * dx);
            return crossVal / Math.Sqrt(lenSq);
        }

        private static double Atan2Safe(double y, double x)
        {
            return Math.Atan2(y, x);
        }

        private static bool IsMiddlePointOnSegmentXY(
            string pPrev,
            string pMid,
            string pNext,
            IReadOnlyDictionary<string, ShellPoint3D> pointCoords,
            double tol)
        {
            var a = pointCoords[pPrev];
            var b = pointCoords[pMid];
            var c = pointCoords[pNext];

            var dx = c.X - a.X;
            var dy = c.Y - a.Y;
            var lenSq = dx * dx + dy * dy;
            if (lenSq <= 0.0)
            {
                return false;
            }

            var distLine = Math.Abs((b.X - a.X) * dy - (b.Y - a.Y) * dx) / Math.Sqrt(lenSq);
            if (distLine > tol)
            {
                return false;
            }

            var t = ((b.X - a.X) * dx + (b.Y - a.Y) * dy) / lenSq;
            return t > 0.0 && t < 1.0;
        }

        private static double Orient2D(double ax, double ay, double bx, double by, double cx, double cy)
        {
            return (bx - ax) * (cy - ay) - (by - ay) * (cx - ax);
        }

        private static string CanonicalLoopKeyGeneric(IReadOnlyList<string> loopPts)
        {
            var s1 = string.Join("|", BuildMinRotationGeneric(loopPts));
            var reversed = loopPts.Reverse().ToArray();
            var s2 = string.Join("|", BuildMinRotationGeneric(reversed));

            return string.Compare(s1, s2, StringComparison.OrdinalIgnoreCase) <= 0 ? s1 : s2;
        }

        private static string[] BuildMinRotationGeneric(IReadOnlyList<string> src)
        {
            var n = src.Count;
            string bestStr = null;
            string[] best = null;

            for (var i = 0; i < n; i++)
            {
                var temp = new string[n];
                for (var k = 0; k < n; k++)
                {
                    temp[k] = src[(i + k) % n];
                }

                var curStr = string.Join("|", temp);
                if (bestStr == null || string.Compare(curStr, bestStr, StringComparison.OrdinalIgnoreCase) < 0)
                {
                    bestStr = curStr;
                    best = temp;
                }
            }

            return best ?? new string[0];
        }

        private class Edge
        {
            public string P1 { get; set; }
            public string P2 { get; set; }
        }

        private class SplitParam
        {
            public double T { get; set; }
            public string NodeId { get; set; }
        }

        private class FaceLoop
        {
            public string Key { get; set; }
            public string[] Loop { get; set; }
            public double Area { get; set; }
        }
    }
}
