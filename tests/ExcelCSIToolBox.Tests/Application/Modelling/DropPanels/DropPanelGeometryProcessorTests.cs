using System;
using System.Collections.Generic;
using System.Linq;
using ExcelCSIToolBox.Application.Modelling.DropPanels;
using FluentAssertions;
using NetTopologySuite.Geometries;
using Xunit;

namespace ExcelCSIToolBox.Tests.Application.Modelling.DropPanels
{
    public class DropPanelGeometryProcessorTests
    {
        [Fact]
        public void BuildPlan_processes_three_columns_and_preserves_source_mapping_across_irregular_slabs()
        {
            DropPanelOptions options = CreateOptions();
            List<DropPanelColumnInfo> columns = new List<DropPanelColumnInfo>
            {
                Column("C1", 5.7, 2.0, 30.0, 0.0),
                Column("C2", 3.8, 3.8, 30.0, 18.0),
                Column("C3", 8.0, 4.2, 30.0, -12.0)
            };
            DropPanelPreparationSnapshot snapshot = new DropPanelPreparationSnapshot();
            snapshot.Areas.Add(Connect(Area(
                "S1",
                "SLAB-A",
                0.0,
                "DEAD",
                "SET-A",
                Point(0, 0), Point(6, 0), Point(6, 6), Point(4, 6), Point(4, 4), Point(0, 4)),
                "C1", "C2"));
            snapshot.Areas.Add(Connect(Area(
                "S2",
                "SLAB-B",
                12.0,
                "LIVE",
                "SET-B",
                Point(6, 0), Point(12, 0), Point(12, 6), Point(6, 6)),
                "C1", "C3"));
            snapshot.Areas.Add(Connect(Area(
                "S3",
                "SLAB-C",
                -20.0,
                "SUPER",
                "SET-C",
                Point(0, 4), Point(4, 4), Point(4, 6), Point(6, 6), Point(6, 10), Point(0, 10)),
                "C2"));
            snapshot.Openings.Add(new DropPanelAreaInfo
            {
                AreaName = "O1",
                StoryName = "L3",
                Elevation = 30.0,
                IsOpening = true,
                Points = new List<DropPanelPoint3D>
                {
                    Point(3.3, 3.1), Point(3.7, 3.1), Point(3.7, 3.5), Point(3.3, 3.5)
                }
            });
            snapshot.Openings.Add(new DropPanelAreaInfo
            {
                AreaName = "O2",
                StoryName = "L3",
                Elevation = 30.0,
                IsOpening = true,
                Points = new List<DropPanelPoint3D>
                {
                    Point(1.0, 8.0), Point(1.5, 8.0), Point(1.5, 8.5), Point(1.0, 8.5)
                }
            });

            Polygon[] openings = snapshot.Openings.Select(opening => ToPolygon(opening.Points)).ToArray();

            DropPanelGeometryProcessor processor = new DropPanelGeometryProcessor();
            var requests = processor.BuildDropRequests(columns, options);
            var result = processor.BuildPlan(columns, snapshot, requests.Data, options);

            result.IsSuccess.Should().BeTrue(result.Message);
            DropPanelOperationPlan plan = result.Data;
            plan.SourceAreas.Select(area => area.AreaName).Should().OnlyHaveUniqueItems();
            plan.SourceAreas.Select(area => area.AreaName).Should().BeEquivalentTo("S1", "S2", "S3");
            plan.Regions.Should().OnlyContain(region => !string.IsNullOrWhiteSpace(region.SourceAreaName));
            plan.Regions.Should().OnlyContain(region => region.Assignment != null);
            plan.Regions.Where(region => region.IsDrop).Should().OnlyContain(region => region.ResultingSectionProperty == "DROP-300");
            plan.Regions.Where(region => !region.IsDrop).Should().OnlyContain(region =>
                region.ResultingSectionProperty == region.Assignment.SectionProperty);

            plan.Regions.Where(region => region.SourceAreaName == "S1")
                .SelectMany(region => region.ColumnNames)
                .Distinct()
                .Should().Contain(new[] { "C1", "C2" });
            plan.Regions.Where(region => region.IsDrop && region.ColumnNames.Contains("C1"))
                .Select(region => region.SourceAreaName)
                .Distinct()
                .Should().Contain(new[] { "S1", "S2" });

            foreach (DropPanelRegion region in plan.Regions)
            {
                region.Assignment.DirectAreaLoads.Should().NotBeEmpty();
                region.Assignment.ShellUniformLoadSetNames.Should().NotBeEmpty();
                region.Assignment.Local3Direction.Z.Should().Be(1.0);
                Polygon polygon = ToPolygon(region.Points);
                polygon.IsValid.Should().BeTrue();
                polygon.NumInteriorRings.Should().Be(0);
                polygon.Area.Should().BeGreaterThanOrEqualTo(options.MinimumPolygonArea);
                openings.Should().OnlyContain(opening => polygon.Intersection(opening).Area <= options.MinimumPolygonArea,
                    "generated regions must not fill an opening, including openings away from the drop rectangle");
            }

        }

        [Fact]
        public void BuildPlan_returns_invalid_geometry_for_a_self_intersecting_source_ring()
        {
            DropPanelOptions options = CreateOptions();
            DropPanelColumnInfo column = Column("C1", 2.0, 2.0, 30.0, 0.0);
            DropPanelPreparationSnapshot snapshot = new DropPanelPreparationSnapshot();
            snapshot.Areas.Add(Connect(Area(
                "S1",
                "SLAB-A",
                0.0,
                "DEAD",
                "SET-A",
                Point(0, 0), Point(4, 4), Point(0, 4), Point(4, 0)),
                "C1"));
            DropPanelGeometryProcessor processor = new DropPanelGeometryProcessor();
            var requests = processor.BuildDropRequests(new[] { column }, options);

            var result = processor.BuildPlan(new[] { column }, snapshot, requests.Data, options);

            result.IsSuccess.Should().BeTrue("geometry validation details should be returned before ETABS is modified");
            result.Data.IsValid.Should().BeFalse();
            result.Data.ValidationMessages.Should().Contain(message => message.Contains("self-intersecting"));
        }

        [Fact]
        public void BuildDropRequests_uses_column_axis_and_user_angle_rotation_modes()
        {
            DropPanelGeometryProcessor processor = new DropPanelGeometryProcessor();
            DropPanelColumnInfo column = Column("C1", 10.0, 20.0, 30.0, 30.0);
            DropPanelOptions options = CreateOptions();
            options.RotationMode = DropPanelRotationMode.FollowColumnLocalAxis;

            var localAxisResult = processor.BuildDropRequests(new[] { column }, options);

            localAxisResult.IsSuccess.Should().BeTrue();
            localAxisResult.Data[0].RotationDegrees.Should().Be(30.0);

            options.RotationMode = DropPanelRotationMode.UserDefinedAngle;
            options.UserDefinedRotationAngle = -17.5;
            var userResult = processor.BuildDropRequests(new[] { column }, options);
            userResult.Data[0].RotationDegrees.Should().Be(-17.5);
        }

        [Fact]
        public void BuildPlan_only_applies_requests_to_areas_connected_to_their_column_heads()
        {
            DropPanelOptions options = CreateOptions();
            List<DropPanelColumnInfo> columns = new List<DropPanelColumnInfo>
            {
                Column("C1", 4.0, 2.0, 30.0, 0.0),
                Column("C2", 7.0, 2.0, 30.0, 0.0)
            };
            DropPanelPreparationSnapshot snapshot = new DropPanelPreparationSnapshot();
            snapshot.Areas.Add(Connect(Area(
                "S1", "SLAB-A", 0.0, "DEAD", "SET-A",
                Point(0, 0), Point(4, 0), Point(4, 4), Point(0, 4)),
                "C1"));
            snapshot.Areas.Add(Connect(Area(
                "S2", "SLAB-B", 0.0, "LIVE", "SET-B",
                Point(4, 0), Point(8, 0), Point(8, 4), Point(4, 4)),
                "C2"));

            DropPanelGeometryProcessor processor = new DropPanelGeometryProcessor();
            var requests = processor.BuildDropRequests(columns, options);
            var result = processor.BuildPlan(columns, snapshot, requests.Data, options);

            result.IsSuccess.Should().BeTrue(result.Message);
            result.Data.IsValid.Should().BeTrue(string.Join(Environment.NewLine, result.Data.ValidationMessages));
            result.Data.Regions
                .Where(region => region.IsDrop && region.SourceAreaName == "S1")
                .SelectMany(region => region.ColumnNames)
                .Distinct()
                .Should().BeEquivalentTo("C1");
            result.Data.Regions
                .Where(region => region.IsDrop && region.SourceAreaName == "S2")
                .SelectMany(region => region.ColumnNames)
                .Distinct()
                .Should().BeEquivalentTo("C2");
        }

        [Fact]
        public void BuildPlan_splits_all_four_shells_connected_to_one_column_head()
        {
            DropPanelOptions options = CreateOptions();
            DropPanelColumnInfo column = Column("C1", 0.0, 0.0, 30.0, 0.0);
            DropPanelPreparationSnapshot snapshot = new DropPanelPreparationSnapshot();
            snapshot.Areas.Add(Connect(Area(
                "S1", "SLAB-1", 0.0, "DEAD", "SET-1",
                Point(-4, -4), Point(0, -4), Point(0, 0), Point(-4, 0)), "C1"));
            snapshot.Areas.Add(Connect(Area(
                "S2", "SLAB-2", 0.0, "LIVE", "SET-2",
                Point(0, -4), Point(4, -4), Point(4, 0), Point(0, 0)), "C1"));
            snapshot.Areas.Add(Connect(Area(
                "S3", "SLAB-3", 0.0, "SUPER", "SET-3",
                Point(-4, 0), Point(0, 0), Point(0, 4), Point(-4, 4)), "C1"));
            snapshot.Areas.Add(Connect(Area(
                "S4", "SLAB-4", 0.0, "WIND", "SET-4",
                Point(0, 0), Point(4, 0), Point(4, 4), Point(0, 4)), "C1"));

            DropPanelGeometryProcessor processor = new DropPanelGeometryProcessor();
            var requests = processor.BuildDropRequests(new[] { column }, options);
            var result = processor.BuildPlan(new[] { column }, snapshot, requests.Data, options);

            result.IsSuccess.Should().BeTrue(result.Message);
            result.Data.IsValid.Should().BeTrue(string.Join(Environment.NewLine, result.Data.ValidationMessages));
            result.Data.SourceAreas.Select(area => area.AreaName)
                .Should().BeEquivalentTo("S1", "S2", "S3", "S4");

            foreach (DropPanelAreaInfo source in snapshot.Areas)
            {
                List<DropPanelRegion> regions = result.Data.Regions
                    .Where(region => region.SourceAreaName == source.AreaName)
                    .ToList();
                regions.Should().Contain(region => region.IsDrop && region.ResultingSectionProperty == "DROP-300");
                regions.Should().Contain(region => !region.IsDrop && region.ResultingSectionProperty == source.SectionProperty);
                regions.Should().OnlyContain(region => ReferenceEquals(region.Assignment, source.Assignment));
                regions.Where(region => region.IsDrop).Sum(region => ToPolygon(region.Points).Area).Should().BeApproximately(1.0, 0.0001);
                regions.Where(region => !region.IsDrop).Sum(region => ToPolygon(region.Points).Area).Should().BeApproximately(15.0, 0.0001);
            }
        }

        private static DropPanelOptions CreateOptions()
        {
            return new DropPanelOptions
            {
                DropProperty = "DROP-300",
                DropSizeX = 2.0,
                DropSizeY = 2.0,
                RotationMode = DropPanelRotationMode.GlobalXY,
                GeometryTolerance = 0.0001,
                ElevationTolerance = 0.01,
                MinimumPolygonArea = 0.00001,
                PreserveMeshAssignments = false
            };
        }

        private static DropPanelColumnInfo Column(string name, double x, double y, double z, double angle)
        {
            return new DropPanelColumnInfo
            {
                FrameName = name,
                BottomPointName = name + "-B",
                TopPointName = name + "-T",
                StoryName = "L3",
                X = x,
                Y = y,
                Z = z,
                SectionProperty = "COL",
                LocalAxisRotationDegrees = angle,
                IsValid = true
            };
        }

        private static DropPanelAreaInfo Area(
            string name,
            string property,
            double angle,
            string loadPattern,
            string loadSet,
            params DropPanelPoint3D[] points)
        {
            DropPanelAreaAssignmentBackup assignment = new DropPanelAreaAssignmentBackup
            {
                SourceAreaName = name,
                SourceAreaLabel = name,
                StoryName = "L3",
                SectionProperty = property,
                LocalAxisAngle = angle,
                Local3Direction = new DropPanelVector3D(0.0, 0.0, 1.0),
                Diaphragm = "D1",
                Modifiers = new[] { 1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0 }
            };
            assignment.DirectAreaLoads.Add(new DropPanelDirectAreaLoad
            {
                LoadPattern = loadPattern,
                LoadType = "Uniform",
                CoordinateSystem = "Global",
                Direction = 6,
                Value = -1.0,
                ReplaceExistingAssignments = true
            });
            assignment.ShellUniformLoadSetNames.Add(loadSet);
            assignment.Groups.Add("FLOOR-GROUP");

            return new DropPanelAreaInfo
            {
                AreaName = name,
                StoryName = "L3",
                SectionProperty = property,
                Elevation = 30.0,
                Points = points.ToList(),
                Assignment = assignment
            };
        }

        private static DropPanelAreaInfo Connect(DropPanelAreaInfo area, params string[] columnNames)
        {
            area.ConnectedColumnNames.AddRange(columnNames);
            return area;
        }

        private static DropPanelPoint3D Point(double x, double y)
        {
            return new DropPanelPoint3D(x, y, 30.0);
        }

        private static Polygon ToPolygon(IReadOnlyList<DropPanelPoint3D> points)
        {
            Coordinate[] coordinates = points
                .Select(point => new Coordinate(point.X, point.Y))
                .Concat(new[] { new Coordinate(points[0].X, points[0].Y) })
                .ToArray();
            return new GeometryFactory().CreatePolygon(coordinates);
        }
    }
}
