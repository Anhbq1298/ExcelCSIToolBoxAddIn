using System;
using System.Collections.Generic;
using ExcelCSIToolBox.Core.Common.Results;

namespace ExcelCSIToolBox.Application.Modelling.OffsetPolylines
{
    public sealed class OffsetPolylineService
    {
        public OperationResult<OffsetPolylineResult> ValidateClosedBoundary(
            IReadOnlyList<SourceLineSegment> sourceSegments,
            OffsetPolylineOptions options)
        {
            ValidationContext context;
            OperationResult validationResult = TryBuildValidationContext(sourceSegments, Normalize(options), out context);
            if (!validationResult.IsSuccess)
            {
                return OperationResult<OffsetPolylineResult>.Failure(validationResult.Message);
            }

            return OperationResult<OffsetPolylineResult>.Success(CreateValidationResult(context, validationResult.Message));
        }

        public OperationResult<OffsetPolylineResult> CalculateOffset(
            IReadOnlyList<SourceLineSegment> sourceSegments,
            double offsetDistance,
            OffsetPolylineOptions options,
            string groupName)
        {
            OffsetPolylineOptions normalizedOptions = Normalize(options);
            if (double.IsNaN(offsetDistance) || double.IsInfinity(offsetDistance))
            {
                return OperationResult<OffsetPolylineResult>.Failure("Offset distance is invalid.");
            }

            if (Math.Abs(offsetDistance) <= normalizedOptions.ZeroLengthTolerance)
            {
                return OperationResult<OffsetPolylineResult>.Failure("Offset distance cannot be zero.");
            }

            ValidationContext context;
            OperationResult validationResult = TryBuildValidationContext(sourceSegments, normalizedOptions, out context);
            if (!validationResult.IsSuccess)
            {
                return OperationResult<OffsetPolylineResult>.Failure(validationResult.Message);
            }

            OperationResult<List<OffsetPoint2D>> offsetVertices2DResult =
                CalculateOffsetVertices(context, offsetDistance, normalizedOptions);
            if (!offsetVertices2DResult.IsSuccess)
            {
                return OperationResult<OffsetPolylineResult>.Failure(offsetVertices2DResult.Message);
            }

            OperationResult resultValidation = ValidateOffsetResult(
                context,
                offsetVertices2DResult.Data,
                offsetDistance,
                normalizedOptions);
            if (!resultValidation.IsSuccess)
            {
                return OperationResult<OffsetPolylineResult>.Failure(resultValidation.Message);
            }

            var resultVertices3D = new List<OffsetPoint3D>();
            foreach (OffsetPoint2D point in offsetVertices2DResult.Data)
            {
                resultVertices3D.Add(Unproject(context.Plane, point));
            }

            string resultType = offsetDistance > 0
                ? "Outer Closed Polyline"
                : "Inner Closed Polyline";

            var resultSegments = new List<OffsetLineSegment>();
            for (int i = 0; i < resultVertices3D.Count; i++)
            {
                OrderedLineSegment source = context.OrderedSegments[i];
                OffsetPoint3D start = resultVertices3D[i];
                OffsetPoint3D end = resultVertices3D[(i + 1) % resultVertices3D.Count];
                resultSegments.Add(new OffsetLineSegment
                {
                    SourceObjectName = source.SourceObjectName,
                    SourceOrderedIndex = source.OrderedIndex,
                    ResultIndex = i + 1,
                    StartX = start.X,
                    StartY = start.Y,
                    StartZ = start.Z,
                    EndX = end.X,
                    EndY = end.Y,
                    EndZ = end.Z,
                    OffsetDistance = offsetDistance,
                    ResultType = resultType,
                    SourceSectionProperty = source.SourceSectionProperty,
                    IsReversedDuringOrdering = source.IsReversedDuringOrdering
                });
            }

            double resultArea = Math.Abs(SignedArea(offsetVertices2DResult.Data));
            var result = CreateValidationResult(context, "Offset preview calculated.");
            result.OffsetDistance = offsetDistance;
            result.OffsetDirection = offsetDistance > 0 ? "Outward" : "Inward";
            result.ResultType = resultType;
            result.ResultSegments = resultSegments;
            result.ResultVertices = resultVertices3D;
            result.ResultArea = resultArea;
            result.GroupName = groupName;
            result.ValidationMessage = resultValidation.Message;
            return OperationResult<OffsetPolylineResult>.Success(result);
        }

        private static OperationResult TryBuildValidationContext(
            IReadOnlyList<SourceLineSegment> sourceSegments,
            OffsetPolylineOptions options,
            out ValidationContext context)
        {
            context = null;

            if (sourceSegments == null || sourceSegments.Count < 3)
            {
                return OperationResult.Failure("At least three line objects must be selected.");
            }

            var nodes = new List<EndpointNode>();
            var edges = new List<GraphEdge>();
            var edgePairs = new HashSet<string>(StringComparer.Ordinal);

            for (int i = 0; i < sourceSegments.Count; i++)
            {
                SourceLineSegment source = sourceSegments[i];
                if (source == null)
                {
                    return OperationResult.Failure("Only straight ETABS line or frame objects are supported.");
                }

                double length = Distance(source.StartPoint, source.EndPoint);
                if (length <= options.ZeroLengthTolerance)
                {
                    return OperationResult.Failure("One or more selected lines have zero length.");
                }

                int startNodeId = FindOrAddNode(nodes, source.StartPoint, options.CoordinateTolerance);
                int endNodeId = FindOrAddNode(nodes, source.EndPoint, options.CoordinateTolerance);
                if (startNodeId == endNodeId)
                {
                    return OperationResult.Failure("One or more selected lines have zero length.");
                }

                string edgeKey = CreateEdgeKey(startNodeId, endNodeId);
                if (edgePairs.Contains(edgeKey))
                {
                    return OperationResult.Failure("The selected boundary contains duplicated segments.");
                }

                edgePairs.Add(edgeKey);

                var edge = new GraphEdge
                {
                    Index = i,
                    Source = source,
                    StartNodeId = startNodeId,
                    EndNodeId = endNodeId
                };
                edges.Add(edge);
                nodes[startNodeId].EdgeIndexes.Add(i);
                nodes[endNodeId].EdgeIndexes.Add(i);
            }

            foreach (EndpointNode node in nodes)
            {
                if (node.EdgeIndexes.Count == 1)
                {
                    return OperationResult.Failure("The selected boundary contains an open endpoint.");
                }

                if (node.EdgeIndexes.Count > 2)
                {
                    return OperationResult.Failure("The selected boundary contains a branch or T-junction.");
                }
            }

            if (nodes.Count != edges.Count)
            {
                return OperationResult.Failure("The selected objects do not form a closed loop.");
            }

            int componentCount = CountConnectedComponents(nodes, edges);
            if (componentCount > 1)
            {
                return OperationResult.Failure("The selected objects contain multiple closed loops.");
            }

            List<OrderedLineSegment> orderedSegments;
            List<int> orderedNodeIds;
            OperationResult orderResult = TryOrderSegments(nodes, edges, out orderedSegments, out orderedNodeIds);
            if (!orderResult.IsSuccess)
            {
                return orderResult;
            }

            List<OffsetPoint3D> vertices = new List<OffsetPoint3D>();
            foreach (int nodeId in orderedNodeIds)
            {
                vertices.Add(nodes[nodeId].Point);
            }

            PlaneBasis plane;
            OperationResult planeResult = TryBuildPlane(vertices, options, out plane);
            if (!planeResult.IsSuccess)
            {
                return planeResult;
            }

            foreach (GraphEdge edge in edges)
            {
                if (DistanceToPlane(plane, edge.Source.StartPoint) > options.PlaneTolerance ||
                    DistanceToPlane(plane, edge.Source.EndPoint) > options.PlaneTolerance)
                {
                    return OperationResult.Failure("The selected lines are not coplanar.");
                }
            }

            var vertices2D = new List<OffsetPoint2D>();
            foreach (OffsetPoint3D point in vertices)
            {
                vertices2D.Add(Project(plane, point));
            }

            double signedArea = SignedArea(vertices2D);
            if (HasPointOnNonIncidentSegment(vertices2D, options.CoordinateTolerance))
            {
                return OperationResult.Failure("The selected boundary contains a branch or T-junction.");
            }

            if (HasSelfIntersections(vertices2D, options.CoordinateTolerance))
            {
                return OperationResult.Failure("The selected boundary is self-intersecting.");
            }

            if (Math.Abs(signedArea) <= options.AreaTolerance)
            {
                return OperationResult.Failure("The resulting polygon has a non-zero area.");
            }

            context = new ValidationContext
            {
                SourceSegments = new List<SourceLineSegment>(sourceSegments),
                Nodes = nodes,
                Edges = edges,
                OrderedSegments = orderedSegments,
                OrderedNodeIds = orderedNodeIds,
                OriginalVertices3D = vertices,
                OriginalVertices2D = vertices2D,
                Plane = plane,
                SignedSourceArea = signedArea
            };

            return OperationResult.Success("The selected lines form one valid closed loop.");
        }

        private static OperationResult<List<OffsetPoint2D>> CalculateOffsetVertices(
            ValidationContext context,
            double offsetDistance,
            OffsetPolylineOptions options)
        {
            int count = context.OriginalVertices2D.Count;
            var directions = new List<OffsetPoint2D>();
            var outwardNormals = new List<OffsetPoint2D>();
            var shiftedStarts = new List<OffsetPoint2D>();
            var resultVertices = new List<OffsetPoint2D>();
            bool isCounterClockwise = context.SignedSourceArea > 0;

            for (int i = 0; i < count; i++)
            {
                OffsetPoint2D start = context.OriginalVertices2D[i];
                OffsetPoint2D end = context.OriginalVertices2D[(i + 1) % count];
                OffsetPoint2D direction = Subtract(end, start);
                double length = Length(direction);
                if (length <= options.ZeroLengthTolerance)
                {
                    return OperationResult<List<OffsetPoint2D>>.Failure("The offset causes one or more segments to collapse.");
                }

                direction = Scale(direction, 1.0 / length);
                OffsetPoint2D outwardNormal = isCounterClockwise
                    ? new OffsetPoint2D(direction.V, -direction.U)
                    : new OffsetPoint2D(-direction.V, direction.U);

                directions.Add(direction);
                outwardNormals.Add(outwardNormal);
                shiftedStarts.Add(Add(start, Scale(outwardNormal, offsetDistance)));
            }

            double miterLimitDistance = Math.Max(options.CoordinateTolerance, Math.Abs(offsetDistance) * options.MiterLimit);
            for (int i = 0; i < count; i++)
            {
                int previous = (i - 1 + count) % count;
                OffsetPoint2D previousStart = shiftedStarts[previous];
                OffsetPoint2D currentStart = shiftedStarts[i];
                OffsetPoint2D previousDirection = directions[previous];
                OffsetPoint2D currentDirection = directions[i];

                OffsetPoint2D intersection;
                if (!TryIntersectOffsetLines(
                    previousStart,
                    previousDirection,
                    currentStart,
                    currentDirection,
                    options,
                    out intersection))
                {
                    double parallelDistance = Math.Abs(Cross(Subtract(currentStart, previousStart), previousDirection));
                    if (parallelDistance <= options.CoordinateTolerance &&
                        Dot(outwardNormals[previous], outwardNormals[i]) > 0.0)
                    {
                        intersection = Add(context.OriginalVertices2D[i], Scale(outwardNormals[i], offsetDistance));
                    }
                    else
                    {
                        return OperationResult<List<OffsetPoint2D>>.Failure(
                            "A valid corner intersection could not be calculated.");
                    }
                }

                if (!IsFinite(intersection))
                {
                    return OperationResult<List<OffsetPoint2D>>.Failure(
                        "The calculated coordinates contain NaN or Infinity.");
                }

                if (Distance(intersection, context.OriginalVertices2D[i]) > miterLimitDistance)
                {
                    return OperationResult<List<OffsetPoint2D>>.Failure(
                        "The offset produces an excessively large corner extension.");
                }

                resultVertices.Add(intersection);
            }

            return OperationResult<List<OffsetPoint2D>>.Success(resultVertices);
        }

        private static OperationResult ValidateOffsetResult(
            ValidationContext context,
            IReadOnlyList<OffsetPoint2D> resultVertices,
            double offsetDistance,
            OffsetPolylineOptions options)
        {
            if (resultVertices == null || resultVertices.Count < 3)
            {
                return OperationResult.Failure("The offset result must contain at least three valid segments.");
            }

            for (int i = 0; i < resultVertices.Count; i++)
            {
                OffsetPoint2D start = resultVertices[i];
                OffsetPoint2D end = resultVertices[(i + 1) % resultVertices.Count];
                if (!IsFinite(start) || !IsFinite(end))
                {
                    return OperationResult.Failure("The calculated coordinates contain NaN or Infinity.");
                }

                if (Distance(start, end) <= options.ZeroLengthTolerance)
                {
                    return OperationResult.Failure("The offset causes one or more segments to collapse.");
                }
            }

            double resultSignedArea = SignedArea(resultVertices);
            if (Math.Abs(resultSignedArea) <= options.AreaTolerance)
            {
                return OperationResult.Failure("The resulting polygon has zero or near-zero area.");
            }

            if ((context.SignedSourceArea > 0 && resultSignedArea < -options.AreaTolerance) ||
                (context.SignedSourceArea < 0 && resultSignedArea > options.AreaTolerance))
            {
                return offsetDistance < 0
                    ? OperationResult.Failure("The entered inward offset is too large for the selected boundary.")
                    : OperationResult.Failure("The offset causes one or more segments to collapse.");
            }

            if (offsetDistance < 0 &&
                Math.Abs(resultSignedArea) >= Math.Abs(context.SignedSourceArea) - options.AreaTolerance)
            {
                return OperationResult.Failure("The entered inward offset is too large for the selected boundary.");
            }

            if (HasSelfIntersections(resultVertices, options.CoordinateTolerance))
            {
                return OperationResult.Failure("The calculated offset boundary is self-intersecting.");
            }

            if (HasPointOnNonIncidentSegment(resultVertices, options.CoordinateTolerance))
            {
                return OperationResult.Failure("The offset produces multiple disconnected boundaries.");
            }

            return OperationResult.Success("Offset result is valid.");
        }

        private static OffsetPolylineResult CreateValidationResult(ValidationContext context, string message)
        {
            return new OffsetPolylineResult
            {
                IsValid = true,
                ValidationMessage = message,
                OffsetDistance = 0,
                OffsetDirection = string.Empty,
                ResultType = string.Empty,
                PolygonOrientation = context.SignedSourceArea >= 0 ? "Counterclockwise" : "Clockwise",
                PlaneOrigin = context.Plane.Origin,
                PlaneNormal = context.Plane.Normal,
                PlaneXAxis = context.Plane.XAxis,
                PlaneYAxis = context.Plane.YAxis,
                OriginalSegments = context.SourceSegments,
                OrderedSegments = context.OrderedSegments,
                ResultSegments = new List<OffsetLineSegment>(),
                OriginalVertices = context.OriginalVertices3D,
                ResultVertices = new List<OffsetPoint3D>(),
                SourceArea = Math.Abs(context.SignedSourceArea),
                ResultArea = 0,
                DetectedVertexCount = context.OriginalVertices3D.Count
            };
        }

        private static OperationResult TryOrderSegments(
            IReadOnlyList<EndpointNode> nodes,
            IReadOnlyList<GraphEdge> edges,
            out List<OrderedLineSegment> orderedSegments,
            out List<int> orderedNodeIds)
        {
            orderedSegments = new List<OrderedLineSegment>();
            orderedNodeIds = new List<int>();

            if (edges.Count == 0)
            {
                return OperationResult.Failure("The selected objects do not form a closed loop.");
            }

            var usedEdges = new HashSet<int>();
            GraphEdge first = edges[0];
            int startNode = first.StartNodeId;
            int currentNode = first.EndNodeId;
            int previousEdgeIndex = first.Index;

            orderedNodeIds.Add(startNode);
            orderedSegments.Add(CreateOrderedSegment(first, startNode, currentNode, 1, nodes));
            usedEdges.Add(first.Index);

            while (currentNode != startNode)
            {
                orderedNodeIds.Add(currentNode);

                int nextEdgeIndex = -1;
                foreach (int candidateEdgeIndex in nodes[currentNode].EdgeIndexes)
                {
                    if (candidateEdgeIndex != previousEdgeIndex)
                    {
                        nextEdgeIndex = candidateEdgeIndex;
                        break;
                    }
                }

                if (nextEdgeIndex < 0 || usedEdges.Contains(nextEdgeIndex))
                {
                    return OperationResult.Failure("The selected objects do not form a closed loop.");
                }

                GraphEdge nextEdge = edges[nextEdgeIndex];
                int nextNode = nextEdge.StartNodeId == currentNode
                    ? nextEdge.EndNodeId
                    : nextEdge.StartNodeId;

                orderedSegments.Add(CreateOrderedSegment(
                    nextEdge,
                    currentNode,
                    nextNode,
                    orderedSegments.Count + 1,
                    nodes));

                usedEdges.Add(nextEdgeIndex);
                previousEdgeIndex = nextEdgeIndex;
                currentNode = nextNode;

                if (orderedSegments.Count > edges.Count)
                {
                    return OperationResult.Failure("The selected objects do not form a closed loop.");
                }
            }

            if (usedEdges.Count != edges.Count || orderedSegments.Count != edges.Count)
            {
                return OperationResult.Failure("The selected objects contain multiple closed loops.");
            }

            return OperationResult.Success();
        }

        private static OrderedLineSegment CreateOrderedSegment(
            GraphEdge edge,
            int orderedStartNode,
            int orderedEndNode,
            int orderedIndex,
            IReadOnlyList<EndpointNode> nodes)
        {
            SourceLineSegment source = edge.Source;
            return new OrderedLineSegment
            {
                SourceObjectName = source.ObjectName,
                SourceSelectionIndex = source.SelectionIndex,
                OrderedIndex = orderedIndex,
                OriginalStartPoint = source.StartPoint,
                OriginalEndPoint = source.EndPoint,
                OrderedStartPoint = nodes[orderedStartNode].Point,
                OrderedEndPoint = nodes[orderedEndNode].Point,
                IsReversedDuringOrdering = orderedStartNode != edge.StartNodeId,
                SourceSectionProperty = source.SectionProperty
            };
        }

        private static OperationResult TryBuildPlane(
            IReadOnlyList<OffsetPoint3D> vertices,
            OffsetPolylineOptions options,
            out PlaneBasis plane)
        {
            plane = null;
            if (vertices == null || vertices.Count < 3)
            {
                return OperationResult.Failure("At least three line objects must be selected.");
            }

            OffsetPoint3D normal = NewellNormal(vertices);
            if (Length(normal) <= options.PlaneTolerance)
            {
                normal = FindFallbackNormal(vertices, options.PlaneTolerance);
            }

            if (Length(normal) <= options.PlaneTolerance)
            {
                return OperationResult.Failure("The resulting polygon has a non-zero area.");
            }

            normal = Normalize(normal);
            OffsetPoint3D origin = vertices[0];
            OffsetPoint3D xAxis = new OffsetPoint3D(0, 0, 0);
            for (int i = 1; i < vertices.Count; i++)
            {
                OffsetPoint3D candidate = Subtract(vertices[i], origin);
                candidate = Subtract(candidate, Scale(normal, Dot(candidate, normal)));
                if (Length(candidate) > options.ZeroLengthTolerance)
                {
                    xAxis = Normalize(candidate);
                    break;
                }
            }

            if (Length(xAxis) <= options.ZeroLengthTolerance)
            {
                return OperationResult.Failure("The resulting polygon has a non-zero area.");
            }

            OffsetPoint3D yAxis = Normalize(Cross(normal, xAxis));
            plane = new PlaneBasis
            {
                Origin = origin,
                Normal = normal,
                XAxis = xAxis,
                YAxis = yAxis
            };

            return OperationResult.Success();
        }

        private static bool TryIntersectOffsetLines(
            OffsetPoint2D a,
            OffsetPoint2D directionA,
            OffsetPoint2D b,
            OffsetPoint2D directionB,
            OffsetPolylineOptions options,
            out OffsetPoint2D intersection)
        {
            intersection = new OffsetPoint2D();
            double denominator = Cross(directionA, directionB);
            if (Math.Abs(denominator) <= options.ParallelTolerance)
            {
                return false;
            }

            OffsetPoint2D delta = Subtract(b, a);
            double t = Cross(delta, directionB) / denominator;
            intersection = Add(a, Scale(directionA, t));
            return true;
        }

        private static int FindOrAddNode(List<EndpointNode> nodes, OffsetPoint3D point, double tolerance)
        {
            for (int i = 0; i < nodes.Count; i++)
            {
                if (Distance(nodes[i].Point, point) <= tolerance)
                {
                    nodes[i].AddPoint(point);
                    return i;
                }
            }

            nodes.Add(new EndpointNode(nodes.Count, point));
            return nodes.Count - 1;
        }

        private static int CountConnectedComponents(IReadOnlyList<EndpointNode> nodes, IReadOnlyList<GraphEdge> edges)
        {
            var visited = new HashSet<int>();
            int count = 0;
            for (int i = 0; i < nodes.Count; i++)
            {
                if (visited.Contains(i))
                {
                    continue;
                }

                count++;
                var queue = new Queue<int>();
                queue.Enqueue(i);
                visited.Add(i);
                while (queue.Count > 0)
                {
                    int nodeId = queue.Dequeue();
                    foreach (int edgeIndex in nodes[nodeId].EdgeIndexes)
                    {
                        GraphEdge edge = edges[edgeIndex];
                        int otherNodeId = edge.StartNodeId == nodeId ? edge.EndNodeId : edge.StartNodeId;
                        if (visited.Add(otherNodeId))
                        {
                            queue.Enqueue(otherNodeId);
                        }
                    }
                }
            }

            return count;
        }

        private static bool HasSelfIntersections(IReadOnlyList<OffsetPoint2D> vertices, double tolerance)
        {
            int count = vertices == null ? 0 : vertices.Count;
            for (int i = 0; i < count; i++)
            {
                OffsetPoint2D a1 = vertices[i];
                OffsetPoint2D a2 = vertices[(i + 1) % count];
                for (int j = i + 1; j < count; j++)
                {
                    if (AreAdjacentSegments(i, j, count))
                    {
                        continue;
                    }

                    OffsetPoint2D b1 = vertices[j];
                    OffsetPoint2D b2 = vertices[(j + 1) % count];
                    if (SegmentsIntersect(a1, a2, b1, b2, tolerance))
                    {
                        return true;
                    }
                }
            }

            return false;
        }

        private static bool HasPointOnNonIncidentSegment(IReadOnlyList<OffsetPoint2D> vertices, double tolerance)
        {
            int count = vertices == null ? 0 : vertices.Count;
            for (int vertexIndex = 0; vertexIndex < count; vertexIndex++)
            {
                OffsetPoint2D point = vertices[vertexIndex];
                for (int segmentIndex = 0; segmentIndex < count; segmentIndex++)
                {
                    int previousSegment = (vertexIndex - 1 + count) % count;
                    if (segmentIndex == vertexIndex || segmentIndex == previousSegment)
                    {
                        continue;
                    }

                    OffsetPoint2D start = vertices[segmentIndex];
                    OffsetPoint2D end = vertices[(segmentIndex + 1) % count];
                    if (PointOnSegment(point, start, end, tolerance))
                    {
                        return true;
                    }
                }
            }

            return false;
        }

        private static bool SegmentsIntersect(
            OffsetPoint2D a,
            OffsetPoint2D b,
            OffsetPoint2D c,
            OffsetPoint2D d,
            double tolerance)
        {
            if (!BoundingBoxesOverlap(a, b, c, d, tolerance))
            {
                return false;
            }

            double o1 = Cross(Subtract(b, a), Subtract(c, a));
            double o2 = Cross(Subtract(b, a), Subtract(d, a));
            double o3 = Cross(Subtract(d, c), Subtract(a, c));
            double o4 = Cross(Subtract(d, c), Subtract(b, c));

            if (Math.Abs(o1) <= tolerance && PointOnSegment(c, a, b, tolerance)) return true;
            if (Math.Abs(o2) <= tolerance && PointOnSegment(d, a, b, tolerance)) return true;
            if (Math.Abs(o3) <= tolerance && PointOnSegment(a, c, d, tolerance)) return true;
            if (Math.Abs(o4) <= tolerance && PointOnSegment(b, c, d, tolerance)) return true;

            return ((o1 > tolerance && o2 < -tolerance) || (o1 < -tolerance && o2 > tolerance)) &&
                   ((o3 > tolerance && o4 < -tolerance) || (o3 < -tolerance && o4 > tolerance));
        }

        private static bool PointOnSegment(OffsetPoint2D point, OffsetPoint2D start, OffsetPoint2D end, double tolerance)
        {
            OffsetPoint2D segment = Subtract(end, start);
            OffsetPoint2D toPoint = Subtract(point, start);
            if (Math.Abs(Cross(segment, toPoint)) > tolerance)
            {
                return false;
            }

            double dot = Dot(toPoint, segment);
            if (dot < tolerance)
            {
                return false;
            }

            double lengthSquared = Dot(segment, segment);
            return dot < lengthSquared - tolerance;
        }

        private static bool BoundingBoxesOverlap(
            OffsetPoint2D a,
            OffsetPoint2D b,
            OffsetPoint2D c,
            OffsetPoint2D d,
            double tolerance)
        {
            return Math.Min(a.U, b.U) <= Math.Max(c.U, d.U) + tolerance &&
                   Math.Max(a.U, b.U) + tolerance >= Math.Min(c.U, d.U) &&
                   Math.Min(a.V, b.V) <= Math.Max(c.V, d.V) + tolerance &&
                   Math.Max(a.V, b.V) + tolerance >= Math.Min(c.V, d.V);
        }

        private static bool AreAdjacentSegments(int left, int right, int count)
        {
            return Math.Abs(left - right) == 1 || (left == 0 && right == count - 1);
        }

        private static double SignedArea(IReadOnlyList<OffsetPoint2D> vertices)
        {
            if (vertices == null || vertices.Count < 3)
            {
                return 0;
            }

            double area = 0;
            for (int i = 0; i < vertices.Count; i++)
            {
                OffsetPoint2D a = vertices[i];
                OffsetPoint2D b = vertices[(i + 1) % vertices.Count];
                area += a.U * b.V - b.U * a.V;
            }

            return area * 0.5;
        }

        private static OffsetPoint2D Project(PlaneBasis plane, OffsetPoint3D point)
        {
            OffsetPoint3D relative = Subtract(point, plane.Origin);
            return new OffsetPoint2D(Dot(relative, plane.XAxis), Dot(relative, plane.YAxis));
        }

        private static OffsetPoint3D Unproject(PlaneBasis plane, OffsetPoint2D point)
        {
            return Add(plane.Origin, Add(Scale(plane.XAxis, point.U), Scale(plane.YAxis, point.V)));
        }

        private static double DistanceToPlane(PlaneBasis plane, OffsetPoint3D point)
        {
            return Math.Abs(Dot(Subtract(point, plane.Origin), plane.Normal));
        }

        private static OffsetPoint3D NewellNormal(IReadOnlyList<OffsetPoint3D> vertices)
        {
            double x = 0;
            double y = 0;
            double z = 0;
            for (int i = 0; i < vertices.Count; i++)
            {
                OffsetPoint3D current = vertices[i];
                OffsetPoint3D next = vertices[(i + 1) % vertices.Count];
                x += (current.Y - next.Y) * (current.Z + next.Z);
                y += (current.Z - next.Z) * (current.X + next.X);
                z += (current.X - next.X) * (current.Y + next.Y);
            }

            return new OffsetPoint3D(x, y, z);
        }

        private static OffsetPoint3D FindFallbackNormal(IReadOnlyList<OffsetPoint3D> vertices, double tolerance)
        {
            for (int i = 0; i < vertices.Count - 2; i++)
            {
                for (int j = i + 1; j < vertices.Count - 1; j++)
                {
                    for (int k = j + 1; k < vertices.Count; k++)
                    {
                        OffsetPoint3D a = Subtract(vertices[j], vertices[i]);
                        OffsetPoint3D b = Subtract(vertices[k], vertices[i]);
                        OffsetPoint3D normal = Cross(a, b);
                        if (Length(normal) > tolerance)
                        {
                            return normal;
                        }
                    }
                }
            }

            return new OffsetPoint3D(0, 0, 0);
        }

        private static OffsetPolylineOptions Normalize(OffsetPolylineOptions options)
        {
            OffsetPolylineOptions source = options ?? new OffsetPolylineOptions();
            return new OffsetPolylineOptions
            {
                CoordinateTolerance = PositiveOrDefault(source.CoordinateTolerance, 0.000001),
                PlaneTolerance = PositiveOrDefault(source.PlaneTolerance, PositiveOrDefault(source.CoordinateTolerance, 0.000001)),
                ZeroLengthTolerance = PositiveOrDefault(source.ZeroLengthTolerance, PositiveOrDefault(source.CoordinateTolerance, 0.000001)),
                ParallelTolerance = PositiveOrDefault(source.ParallelTolerance, 0.000000001),
                AreaTolerance = PositiveOrDefault(source.AreaTolerance, 0.000000001),
                MiterLimit = source.MiterLimit < 1.0 || double.IsNaN(source.MiterLimit) || double.IsInfinity(source.MiterLimit)
                    ? 10.0
                    : source.MiterLimit
            };
        }

        private static double PositiveOrDefault(double value, double fallback)
        {
            return value > 0 && !double.IsNaN(value) && !double.IsInfinity(value) ? value : fallback;
        }

        private static string CreateEdgeKey(int left, int right)
        {
            return left < right
                ? left.ToString(System.Globalization.CultureInfo.InvariantCulture) + ":" + right.ToString(System.Globalization.CultureInfo.InvariantCulture)
                : right.ToString(System.Globalization.CultureInfo.InvariantCulture) + ":" + left.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }

        private static bool IsFinite(OffsetPoint2D point)
        {
            return !double.IsNaN(point.U) &&
                   !double.IsNaN(point.V) &&
                   !double.IsInfinity(point.U) &&
                   !double.IsInfinity(point.V);
        }

        private static OffsetPoint2D Add(OffsetPoint2D left, OffsetPoint2D right)
        {
            return new OffsetPoint2D(left.U + right.U, left.V + right.V);
        }

        private static OffsetPoint3D Add(OffsetPoint3D left, OffsetPoint3D right)
        {
            return new OffsetPoint3D(left.X + right.X, left.Y + right.Y, left.Z + right.Z);
        }

        private static OffsetPoint2D Subtract(OffsetPoint2D left, OffsetPoint2D right)
        {
            return new OffsetPoint2D(left.U - right.U, left.V - right.V);
        }

        private static OffsetPoint3D Subtract(OffsetPoint3D left, OffsetPoint3D right)
        {
            return new OffsetPoint3D(left.X - right.X, left.Y - right.Y, left.Z - right.Z);
        }

        private static OffsetPoint2D Scale(OffsetPoint2D point, double scale)
        {
            return new OffsetPoint2D(point.U * scale, point.V * scale);
        }

        private static OffsetPoint3D Scale(OffsetPoint3D point, double scale)
        {
            return new OffsetPoint3D(point.X * scale, point.Y * scale, point.Z * scale);
        }

        private static double Dot(OffsetPoint2D left, OffsetPoint2D right)
        {
            return left.U * right.U + left.V * right.V;
        }

        private static double Dot(OffsetPoint3D left, OffsetPoint3D right)
        {
            return left.X * right.X + left.Y * right.Y + left.Z * right.Z;
        }

        private static double Cross(OffsetPoint2D left, OffsetPoint2D right)
        {
            return left.U * right.V - left.V * right.U;
        }

        private static OffsetPoint3D Cross(OffsetPoint3D left, OffsetPoint3D right)
        {
            return new OffsetPoint3D(
                left.Y * right.Z - left.Z * right.Y,
                left.Z * right.X - left.X * right.Z,
                left.X * right.Y - left.Y * right.X);
        }

        private static double Length(OffsetPoint2D point)
        {
            return Math.Sqrt(Dot(point, point));
        }

        private static double Length(OffsetPoint3D point)
        {
            return Math.Sqrt(Dot(point, point));
        }

        private static double Distance(OffsetPoint2D left, OffsetPoint2D right)
        {
            return Length(Subtract(left, right));
        }

        private static double Distance(OffsetPoint3D left, OffsetPoint3D right)
        {
            return Length(Subtract(left, right));
        }

        private static OffsetPoint3D Normalize(OffsetPoint3D point)
        {
            double length = Length(point);
            return length <= 0 ? new OffsetPoint3D(0, 0, 0) : Scale(point, 1.0 / length);
        }

        private sealed class ValidationContext
        {
            public IReadOnlyList<SourceLineSegment> SourceSegments { get; set; }
            public IReadOnlyList<EndpointNode> Nodes { get; set; }
            public IReadOnlyList<GraphEdge> Edges { get; set; }
            public IReadOnlyList<OrderedLineSegment> OrderedSegments { get; set; }
            public IReadOnlyList<int> OrderedNodeIds { get; set; }
            public IReadOnlyList<OffsetPoint3D> OriginalVertices3D { get; set; }
            public IReadOnlyList<OffsetPoint2D> OriginalVertices2D { get; set; }
            public PlaneBasis Plane { get; set; }
            public double SignedSourceArea { get; set; }
        }

        private sealed class PlaneBasis
        {
            public OffsetPoint3D Origin { get; set; }
            public OffsetPoint3D Normal { get; set; }
            public OffsetPoint3D XAxis { get; set; }
            public OffsetPoint3D YAxis { get; set; }
        }

        private sealed class EndpointNode
        {
            private int _pointCount;

            public EndpointNode(int id, OffsetPoint3D point)
            {
                Id = id;
                Point = point;
                EdgeIndexes = new List<int>();
                _pointCount = 1;
            }

            public int Id { get; private set; }
            public OffsetPoint3D Point { get; private set; }
            public List<int> EdgeIndexes { get; private set; }

            public void AddPoint(OffsetPoint3D point)
            {
                _pointCount++;
                double count = _pointCount;
                Point = new OffsetPoint3D(
                    (Point.X * (count - 1.0) + point.X) / count,
                    (Point.Y * (count - 1.0) + point.Y) / count,
                    (Point.Z * (count - 1.0) + point.Z) / count);
            }
        }

        private sealed class GraphEdge
        {
            public int Index { get; set; }
            public SourceLineSegment Source { get; set; }
            public int StartNodeId { get; set; }
            public int EndNodeId { get; set; }
        }
    }
}
