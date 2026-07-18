using System;
using System.Collections.Generic;
using System.Linq;
using ExcelCSIToolBox.Core.Common.Results;
using NetTopologySuite.Geometries;
using NetTopologySuite.Geometries.Utilities;
using NetTopologySuite.Operation.Polygonize;
using NetTopologySuite.Operation.Union;
using NetTopologySuite.Triangulate.Polygon;

namespace ExcelCSIToolBox.Application.Modelling.DropPanels
{
    public sealed class DropPanelGeometryProcessor
    {
        public OperationResult<IReadOnlyList<DropPanelRequest>> BuildDropRequests(
            IReadOnlyList<DropPanelColumnInfo> columns,
            DropPanelOptions options)
        {
            OperationResult validation = ValidateInputs(columns, options);
            if (!validation.IsSuccess)
            {
                return OperationResult<IReadOnlyList<DropPanelRequest>>.Failure(validation.Message);
            }

            List<DropPanelRequest> requests = new List<DropPanelRequest>();
            foreach (DropPanelColumnInfo column in columns.Where(item => item != null && item.IsValid))
            {
                double angle = GetRotationAngle(column, options);
                double radians = angle * Math.PI / 180.0;
                double cosine = Math.Cos(radians);
                double sine = Math.Sin(radians);
                double halfX = options.DropSizeX / 2.0;
                double halfY = options.DropSizeY / 2.0;
                double[,] localCorners =
                {
                    { -halfX, -halfY },
                    { halfX, -halfY },
                    { halfX, halfY },
                    { -halfX, halfY }
                };

                DropPanelRequest request = new DropPanelRequest
                {
                    ColumnName = column.FrameName,
                    StoryName = column.StoryName,
                    Elevation = column.Z,
                    RotationDegrees = angle,
                    DropProperty = options.DropProperty
                };

                for (int index = 0; index < 4; index++)
                {
                    double localX = localCorners[index, 0];
                    double localY = localCorners[index, 1];
                    request.Points.Add(new DropPanelPoint3D(
                        column.X + localX * cosine - localY * sine,
                        column.Y + localX * sine + localY * cosine,
                        column.Z));
                }

                requests.Add(request);
            }

            return OperationResult<IReadOnlyList<DropPanelRequest>>.Success(requests);
        }

        public OperationResult<DropPanelOperationPlan> BuildPlan(
            IReadOnlyList<DropPanelColumnInfo> columns,
            DropPanelPreparationSnapshot snapshot,
            IReadOnlyList<DropPanelRequest> requests,
            DropPanelOptions options)
        {
            if (snapshot == null)
            {
                return OperationResult<DropPanelOperationPlan>.Failure("The ETABS area snapshot is required.");
            }

            if (requests == null || requests.Count == 0)
            {
                return OperationResult<DropPanelOperationPlan>.Failure("At least one drop request is required.");
            }

            GeometryFactory factory = CreateGeometryFactory(options.GeometryTolerance);
            DropPanelOperationPlan plan = new DropPanelOperationPlan();
            plan.Columns.AddRange(columns ?? new DropPanelColumnInfo[0]);
            plan.Requests.AddRange(requests);
            plan.Openings.AddRange(snapshot.Openings ?? new List<DropPanelAreaInfo>());
            plan.ModelPath = snapshot.ModelPath;
            plan.PresentUnits = snapshot.PresentUnits;

            Dictionary<string, Geometry> requestGeometries = new Dictionary<string, Geometry>(StringComparer.OrdinalIgnoreCase);
            Dictionary<string, bool> requestHits = new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);
            foreach (DropPanelRequest request in requests)
            {
                Geometry requestGeometry;
                OperationResult geometryResult = TryCreatePolygon(factory, request.Points, options.GeometryTolerance, out requestGeometry);
                if (!geometryResult.IsSuccess)
                {
                    plan.ValidationMessages.Add("Column '" + request.ColumnName + "': " + geometryResult.Message);
                    continue;
                }

                requestGeometries[request.ColumnName] = requestGeometry;
                requestHits[request.ColumnName] = false;
            }

            foreach (DropPanelAreaInfo sourceArea in snapshot.Areas ?? new List<DropPanelAreaInfo>())
            {
                if (sourceArea == null || sourceArea.IsOpening)
                {
                    continue;
                }

                Geometry sourceGeometry;
                OperationResult sourceResult = TryCreatePolygon(factory, sourceArea.Points, options.GeometryTolerance, out sourceGeometry);
                if (!sourceResult.IsSuccess)
                {
                    plan.ValidationMessages.Add("Source area '" + sourceArea.AreaName + "': " + sourceResult.Message);
                    continue;
                }

                List<KeyValuePair<DropPanelRequest, Geometry>> relevantRequests = new List<KeyValuePair<DropPanelRequest, Geometry>>();
                foreach (DropPanelRequest request in requests)
                {
                    Geometry requestGeometry;
                    if (!requestGeometries.TryGetValue(request.ColumnName, out requestGeometry) ||
                        Math.Abs(request.Elevation - sourceArea.Elevation) > options.ElevationTolerance ||
                        !sourceGeometry.EnvelopeInternal.Intersects(requestGeometry.EnvelopeInternal))
                    {
                        continue;
                    }

                    Geometry intersection = SafeIntersection(sourceGeometry, requestGeometry);
                    if (!intersection.IsEmpty && intersection.Area >= options.MinimumPolygonArea)
                    {
                        relevantRequests.Add(new KeyValuePair<DropPanelRequest, Geometry>(request, requestGeometry));
                        requestHits[request.ColumnName] = true;
                    }
                }

                if (relevantRequests.Count == 0)
                {
                    continue;
                }

                // Keep affected source geometry in an invalid plan so Preview can explain the problem visually.
                plan.SourceAreas.Add(sourceArea);

                if (sourceArea.Assignment == null)
                {
                    plan.ValidationMessages.Add("Assignments were not backed up for source area '" + sourceArea.AreaName + "'.");
                    continue;
                }

                if (options.PreserveLocalAxes && sourceArea.Assignment.UsesAdvancedLocalAxes)
                {
                    plan.ValidationMessages.Add(
                        "Source area '" + sourceArea.AreaName + "' uses advanced local axes that cannot be restored by the referenced ETABS API.");
                    continue;
                }

                List<Geometry> dropInputs = relevantRequests.Select(item => item.Value).ToList();
                Geometry dropUnion = GeometryFixer.Fix(UnaryUnionOp.Union(dropInputs));
                Geometry openingUnion = BuildOpeningUnion(factory, snapshot.Openings, sourceArea.Elevation, sourceGeometry, options);
                Geometry dropPart = SafeIntersection(sourceGeometry, dropUnion);
                if (!openingUnion.IsEmpty)
                {
                    dropPart = SafeDifference(dropPart, openingUnion);
                }

                Geometry normalPart = SafeDifference(sourceGeometry, dropPart);
                if (!openingUnion.IsEmpty)
                {
                    normalPart = SafeDifference(normalPart, openingUnion);
                }

                List<Polygon> dropPolygons = ExtractSimplePolygons(dropPart, factory, options);
                List<Polygon> normalPolygons = ExtractSimplePolygons(normalPart, factory, options);
                if (dropPolygons.Count == 0)
                {
                    plan.ValidationMessages.Add("The requested drop does not create a valid region in source area '" + sourceArea.AreaName + "'.");
                    continue;
                }

                DropPanelAssignmentSignatureBuilder signatureBuilder = new DropPanelAssignmentSignatureBuilder();
                string normalSignature = signatureBuilder.Build(sourceArea.SectionProperty, sourceArea.Assignment);
                string dropSignature = signatureBuilder.Build(options.DropProperty, sourceArea.Assignment);
                AddRegions(plan, sourceArea, normalPolygons, false, relevantRequests, options, normalSignature);
                AddRegions(plan, sourceArea, dropPolygons, true, relevantRequests, options, dropSignature);
            }

            foreach (DropPanelRequest request in requests)
            {
                bool hit;
                if (!requestHits.TryGetValue(request.ColumnName, out hit) || !hit)
                {
                    plan.ValidationMessages.Add("Drop polygon for column '" + request.ColumnName + "' does not intersect a slab area.");
                }
            }

            if (plan.SourceAreas.Count == 0 && plan.ValidationMessages.Count == 0)
            {
                plan.ValidationMessages.Add("No affected slab areas were found at the selected column heads.");
            }

            if (plan.Regions.Count == 0 && plan.ValidationMessages.Count == 0)
            {
                plan.ValidationMessages.Add("The batch geometry operation produced no ETABS regions.");
            }

            return OperationResult<DropPanelOperationPlan>.Success(
                plan,
                plan.IsValid ? "Drop panel preview prepared." : "Drop panel preview contains validation errors.");
        }

        private static void AddRegions(
            DropPanelOperationPlan plan,
            DropPanelAreaInfo sourceArea,
            IReadOnlyList<Polygon> polygons,
            bool isDrop,
            IReadOnlyList<KeyValuePair<DropPanelRequest, Geometry>> relevantRequests,
            DropPanelOptions options,
            string assignmentSignature)
        {
            List<Polygon> compatiblePolygons = polygons.ToList();
            if (options.MergeAdjacentRegionsOnlyWhenAssignmentSignaturesMatch && !string.IsNullOrWhiteSpace(assignmentSignature))
            {
                // This list contains one source, one result property, and one complete assignment signature.
                MergeCompatiblePolygons(compatiblePolygons, options.GeometryTolerance);
            }

            foreach (Polygon polygon in compatiblePolygons)
            {
                List<DropPanelPoint3D> points = ToPoints(polygon, sourceArea.Elevation, options.GeometryTolerance);
                if (points.Count < 3)
                {
                    plan.ValidationMessages.Add("A generated region for source area '" + sourceArea.AreaName + "' has fewer than three vertices.");
                    continue;
                }

                DropPanelRegion region = new DropPanelRegion
                {
                    SourceAreaName = sourceArea.AreaName,
                    IsDrop = isDrop,
                    ResultingSectionProperty = isDrop ? options.DropProperty : sourceArea.SectionProperty,
                    AssignmentSignature = assignmentSignature,
                    Points = points,
                    Assignment = sourceArea.Assignment
                };

                foreach (KeyValuePair<DropPanelRequest, Geometry> request in relevantRequests)
                {
                    if (!SafeIntersection(polygon, request.Value).IsEmpty)
                    {
                        region.ColumnNames.Add(request.Key.ColumnName);
                    }
                }

                plan.Regions.Add(region);
            }
        }

        private static Geometry BuildOpeningUnion(
            GeometryFactory factory,
            IReadOnlyList<DropPanelAreaInfo> openings,
            double elevation,
            Geometry sourceGeometry,
            DropPanelOptions options)
        {
            List<Geometry> matching = new List<Geometry>();
            foreach (DropPanelAreaInfo opening in openings ?? new DropPanelAreaInfo[0])
            {
                if (opening == null || Math.Abs(opening.Elevation - elevation) > options.ElevationTolerance)
                {
                    continue;
                }

                Geometry geometry;
                OperationResult result = TryCreatePolygon(factory, opening.Points, options.GeometryTolerance, out geometry);
                if (result.IsSuccess && sourceGeometry.EnvelopeInternal.Intersects(geometry.EnvelopeInternal))
                {
                    matching.Add(geometry);
                }
            }

            return matching.Count == 0
                ? factory.CreatePolygon()
                : GeometryFixer.Fix(UnaryUnionOp.Union(matching));
        }

        private static List<Polygon> ExtractSimplePolygons(
            Geometry geometry,
            GeometryFactory factory,
            DropPanelOptions options)
        {
            List<Polygon> result = new List<Polygon>();
            Geometry fixedGeometry = GeometryFixer.Fix(geometry);
            List<Polygon> candidates = EnumeratePolygons(fixedGeometry).ToList();
            if (candidates.Count == 0 && fixedGeometry != null && !fixedGeometry.IsEmpty)
            {
                Polygonizer polygonizer = new Polygonizer();
                polygonizer.Add(fixedGeometry.Boundary);
                candidates.AddRange(polygonizer.GetPolygons().OfType<Polygon>());
            }

            foreach (Polygon polygon in candidates)
            {
                if (polygon.Area < options.MinimumPolygonArea)
                {
                    continue;
                }

                if (polygon.NumInteriorRings == 0)
                {
                    result.Add(polygon);
                    continue;
                }

                Geometry triangles = PolygonTriangulator.Triangulate(polygon);
                foreach (Polygon triangle in EnumeratePolygons(triangles))
                {
                    if (triangle.Area >= options.MinimumPolygonArea && polygon.Covers(triangle))
                    {
                        result.Add(triangle);
                    }
                }
            }

            return result
                .Where(item => item != null && item.IsValid && item.NumInteriorRings == 0 && item.Area >= options.MinimumPolygonArea)
                .OrderBy(item => item.Centroid.Y)
                .ThenBy(item => item.Centroid.X)
                .ThenBy(item => item.Area)
                .ToList();
        }

        private static void MergeCompatiblePolygons(List<Polygon> polygons, double tolerance)
        {
            bool changed = true;
            while (changed)
            {
                changed = false;
                for (int leftIndex = 0; leftIndex < polygons.Count && !changed; leftIndex++)
                {
                    for (int rightIndex = leftIndex + 1; rightIndex < polygons.Count; rightIndex++)
                    {
                        Polygon left = polygons[leftIndex];
                        Polygon right = polygons[rightIndex];
                        if (!left.EnvelopeInternal.Intersects(right.EnvelopeInternal) || !left.Touches(right))
                        {
                            continue;
                        }

                        Geometry union = GeometryFixer.Fix(left.Union(right));
                        Polygon merged = union as Polygon;
                        if (merged == null || merged.NumInteriorRings != 0 || !merged.IsValid ||
                            Math.Abs(merged.Area - left.Area - right.Area) > Math.Max(tolerance * tolerance, 1e-10))
                        {
                            continue;
                        }

                        polygons[leftIndex] = merged;
                        polygons.RemoveAt(rightIndex);
                        changed = true;
                        break;
                    }
                }
            }
        }

        private static IEnumerable<Polygon> EnumeratePolygons(Geometry geometry)
        {
            if (geometry == null || geometry.IsEmpty)
            {
                yield break;
            }

            Polygon polygon = geometry as Polygon;
            if (polygon != null)
            {
                yield return polygon;
                yield break;
            }

            for (int index = 0; index < geometry.NumGeometries; index++)
            {
                foreach (Polygon child in EnumeratePolygons(geometry.GetGeometryN(index)))
                {
                    yield return child;
                }
            }
        }

        private static OperationResult TryCreatePolygon(
            GeometryFactory factory,
            IReadOnlyList<DropPanelPoint3D> points,
            double tolerance,
            out Geometry geometry)
        {
            geometry = null;
            List<Coordinate> coordinates = NormalizeCoordinates(points, tolerance);
            if (coordinates.Count < 4)
            {
                return OperationResult.Failure("The polygon has fewer than three unique vertices.");
            }

            try
            {
                Polygon polygon = factory.CreatePolygon(coordinates.ToArray());
                if (!polygon.IsValid)
                {
                    return OperationResult.Failure("The polygon is self-intersecting or otherwise invalid.");
                }

                geometry = GeometryFixer.Fix(polygon);
                if (geometry == null || geometry.IsEmpty || geometry.Dimension != Dimension.Surface)
                {
                    return OperationResult.Failure("The polygon is empty or collapsed after validation.");
                }

                if (!geometry.IsValid)
                {
                    return OperationResult.Failure("The polygon is self-intersecting or otherwise invalid.");
                }

                return OperationResult.Success();
            }
            catch (Exception ex)
            {
                return OperationResult.Failure("The polygon could not be created: " + ex.Message);
            }
        }

        private static List<Coordinate> NormalizeCoordinates(IReadOnlyList<DropPanelPoint3D> points, double tolerance)
        {
            List<Coordinate> coordinates = new List<Coordinate>();
            foreach (DropPanelPoint3D point in points ?? new DropPanelPoint3D[0])
            {
                if (point == null || double.IsNaN(point.X) || double.IsNaN(point.Y) ||
                    double.IsInfinity(point.X) || double.IsInfinity(point.Y))
                {
                    continue;
                }

                Coordinate coordinate = new Coordinate(point.X, point.Y);
                if (coordinates.Count == 0 || coordinates[coordinates.Count - 1].Distance(coordinate) > tolerance)
                {
                    coordinates.Add(coordinate);
                }
            }

            if (coordinates.Count > 1 && coordinates[0].Distance(coordinates[coordinates.Count - 1]) <= tolerance)
            {
                coordinates.RemoveAt(coordinates.Count - 1);
            }

            RemoveCollinearCoordinates(coordinates, tolerance);
            if (coordinates.Count >= 3)
            {
                coordinates.Add(new Coordinate(coordinates[0]));
            }

            return coordinates;
        }

        private static void RemoveCollinearCoordinates(List<Coordinate> coordinates, double tolerance)
        {
            bool removed = true;
            while (removed && coordinates.Count > 3)
            {
                removed = false;
                for (int index = 0; index < coordinates.Count; index++)
                {
                    Coordinate previous = coordinates[(index + coordinates.Count - 1) % coordinates.Count];
                    Coordinate current = coordinates[index];
                    Coordinate next = coordinates[(index + 1) % coordinates.Count];
                    double cross = (current.X - previous.X) * (next.Y - current.Y) -
                                   (current.Y - previous.Y) * (next.X - current.X);
                    double scale = Math.Max(1.0, previous.Distance(current) + current.Distance(next));
                    if (Math.Abs(cross) <= tolerance * scale)
                    {
                        coordinates.RemoveAt(index);
                        removed = true;
                        break;
                    }
                }
            }
        }

        private static List<DropPanelPoint3D> ToPoints(Polygon polygon, double elevation, double tolerance)
        {
            List<DropPanelPoint3D> points = new List<DropPanelPoint3D>();
            Coordinate[] coordinates = polygon.ExteriorRing.Coordinates;
            int count = coordinates.Length > 1 && coordinates[0].Distance(coordinates[coordinates.Length - 1]) <= tolerance
                ? coordinates.Length - 1
                : coordinates.Length;
            for (int index = 0; index < count; index++)
            {
                points.Add(new DropPanelPoint3D(coordinates[index].X, coordinates[index].Y, elevation));
            }

            return points;
        }

        private static Geometry SafeIntersection(Geometry left, Geometry right)
        {
            try
            {
                return GeometryFixer.Fix(left).Intersection(GeometryFixer.Fix(right));
            }
            catch
            {
                return GeometryFixer.Fix(GeometryFixer.Fix(left).Buffer(0.0))
                    .Intersection(GeometryFixer.Fix(right).Buffer(0.0));
            }
        }

        private static Geometry SafeDifference(Geometry left, Geometry right)
        {
            try
            {
                return GeometryFixer.Fix(left).Difference(GeometryFixer.Fix(right));
            }
            catch
            {
                return GeometryFixer.Fix(GeometryFixer.Fix(left).Buffer(0.0))
                    .Difference(GeometryFixer.Fix(right).Buffer(0.0));
            }
        }

        private static GeometryFactory CreateGeometryFactory(double tolerance)
        {
            double scale = tolerance > 0.0 ? 1.0 / tolerance : 1000000.0;
            return new GeometryFactory(new PrecisionModel(scale));
        }

        private static double GetRotationAngle(DropPanelColumnInfo column, DropPanelOptions options)
        {
            switch (options.RotationMode)
            {
                case DropPanelRotationMode.FollowColumnLocalAxis:
                    return column.LocalAxisRotationDegrees;
                case DropPanelRotationMode.UserDefinedAngle:
                    return options.UserDefinedRotationAngle;
                default:
                    return 0.0;
            }
        }

        private static OperationResult ValidateInputs(IReadOnlyList<DropPanelColumnInfo> columns, DropPanelOptions options)
        {
            if (columns == null || columns.Count == 0 || !columns.Any(item => item != null && item.IsValid))
            {
                return OperationResult.Failure("No valid ETABS columns are available.");
            }

            if (options == null)
            {
                return OperationResult.Failure("Drop panel options are required.");
            }

            if (string.IsNullOrWhiteSpace(options.DropProperty))
            {
                return OperationResult.Failure("Select a drop property.");
            }

            if (options.DropSizeX <= 0.0 || options.DropSizeY <= 0.0)
            {
                return OperationResult.Failure("Drop sizes must be greater than zero.");
            }

            if (options.GeometryTolerance <= 0.0 || options.ElevationTolerance < 0.0 || options.MinimumPolygonArea <= 0.0)
            {
                return OperationResult.Failure("Geometry tolerance and minimum polygon area must be greater than zero.");
            }

            return OperationResult.Success();
        }
    }
}
