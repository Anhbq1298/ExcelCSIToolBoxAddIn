using System.Collections.Generic;
using System.Linq;
using ExcelCSIToolBox.Application.Modelling.OffsetPolylines;
using FluentAssertions;
using Xunit;

namespace ExcelCSIToolBox.Tests
{
    public class OffsetPolylineServiceTests
    {
        private readonly OffsetPolylineService _service = new OffsetPolylineService();

        [Fact]
        public void ValidateClosedBoundary_orders_unordered_mixed_direction_rectangle()
        {
            var segments = new[]
            {
                Segment("R", 10, 0, 0, 10, 5, 0, 3),
                Segment("B", 10, 0, 0, 0, 0, 0, 1),
                Segment("T", 10, 5, 0, 0, 5, 0, 4),
                Segment("L", 0, 0, 0, 0, 5, 0, 2)
            };

            var result = _service.ValidateClosedBoundary(segments, Options());

            result.IsSuccess.Should().BeTrue(result.Message);
            result.Data.OrderedSegments.Should().HaveCount(4);
            result.Data.DetectedVertexCount.Should().Be(4);
            result.Data.OrderedSegments.Any(x => x.IsReversedDuringOrdering).Should().BeTrue();
        }

        [Fact]
        public void CalculateOffset_positive_value_creates_larger_outer_rectangle()
        {
            var result = _service.CalculateOffset(Rectangle(), 1.0, Options(), null);

            result.IsSuccess.Should().BeTrue(result.Message);
            result.Data.ResultType.Should().Be("Outer Closed Polyline");
            result.Data.ResultArea.Should().BeGreaterThan(result.Data.SourceArea);
            Bounds(result.Data.ResultVertices).Should().BeEquivalentTo(new Bounds2D(-1, 11, -1, 6));
        }

        [Fact]
        public void CalculateOffset_negative_value_creates_smaller_inner_rectangle()
        {
            var result = _service.CalculateOffset(Rectangle(), -1.0, Options(), null);

            result.IsSuccess.Should().BeTrue(result.Message);
            result.Data.ResultType.Should().Be("Inner Closed Polyline");
            result.Data.ResultArea.Should().BeLessThan(result.Data.SourceArea);
            Bounds(result.Data.ResultVertices).Should().BeEquivalentTo(new Bounds2D(1, 9, 1, 4));
        }

        [Fact]
        public void ValidateClosedBoundary_rejects_open_boundary()
        {
            var segments = new[]
            {
                Segment("A", 0, 0, 0, 10, 0, 0, 1),
                Segment("B", 10, 0, 0, 10, 5, 0, 2),
                Segment("C", 10, 5, 0, 0, 5, 0, 3)
            };

            var result = _service.ValidateClosedBoundary(segments, Options());

            result.IsSuccess.Should().BeFalse();
            result.Message.Should().Contain("open endpoint");
        }

        [Fact]
        public void ValidateClosedBoundary_rejects_duplicated_segment()
        {
            var segments = new[]
            {
                Segment("A", 0, 0, 0, 10, 0, 0, 1),
                Segment("B", 10, 0, 0, 10, 5, 0, 2),
                Segment("C", 10, 5, 0, 0, 5, 0, 3),
                Segment("D", 0, 5, 0, 0, 0, 0, 4),
                Segment("A2", 10, 0, 0, 0, 0, 0, 5)
            };

            var result = _service.ValidateClosedBoundary(segments, Options());

            result.IsSuccess.Should().BeFalse();
            result.Message.Should().Contain("duplicated");
        }

        [Fact]
        public void ValidateClosedBoundary_rejects_self_intersecting_source()
        {
            var segments = new[]
            {
                Segment("A", 0, 0, 0, 10, 10, 0, 1),
                Segment("B", 10, 10, 0, 0, 10, 0, 2),
                Segment("C", 0, 10, 0, 10, 0, 0, 3),
                Segment("D", 10, 0, 0, 0, 0, 0, 4)
            };

            var result = _service.ValidateClosedBoundary(segments, Options());

            result.IsSuccess.Should().BeFalse();
            result.Message.Should().Contain("self-intersecting");
        }

        [Fact]
        public void CalculateOffset_preserves_segment_count_with_consecutive_collinear_segments()
        {
            var segments = new[]
            {
                Segment("A", 0, 0, 0, 5, 0, 0, 1),
                Segment("B", 5, 0, 0, 10, 0, 0, 2),
                Segment("C", 10, 0, 0, 10, 5, 0, 3),
                Segment("D", 10, 5, 0, 0, 5, 0, 4),
                Segment("E", 0, 5, 0, 0, 0, 0, 5)
            };

            var result = _service.CalculateOffset(segments, 1.0, Options(), null);

            result.IsSuccess.Should().BeTrue(result.Message);
            result.Data.ResultSegments.Should().HaveCount(5);
            result.Data.ResultSegments[0].SourceObjectName.Should().Be("A");
            result.Data.ResultSegments[1].SourceObjectName.Should().Be("B");
        }

        [Fact]
        public void CalculateOffset_handles_vertical_polygon_plane()
        {
            var segments = new[]
            {
                Segment("A", 0, 3, 0, 10, 3, 0, 1),
                Segment("B", 10, 3, 0, 10, 3, 5, 2),
                Segment("C", 10, 3, 5, 0, 3, 5, 3),
                Segment("D", 0, 3, 5, 0, 3, 0, 4)
            };

            var result = _service.CalculateOffset(segments, 1.0, Options(), null);

            result.IsSuccess.Should().BeTrue(result.Message);
            result.Data.ResultVertices.Should().OnlyContain(p => System.Math.Abs(p.Y - 3) < 0.000001);
            result.Data.ResultSegments.Should().HaveCount(4);
        }

        private static IReadOnlyList<SourceLineSegment> Rectangle()
        {
            return new[]
            {
                Segment("A", 0, 0, 0, 10, 0, 0, 1),
                Segment("B", 10, 0, 0, 10, 5, 0, 2),
                Segment("C", 10, 5, 0, 0, 5, 0, 3),
                Segment("D", 0, 5, 0, 0, 0, 0, 4)
            };
        }

        private static SourceLineSegment Segment(
            string name,
            double sx,
            double sy,
            double sz,
            double ex,
            double ey,
            double ez,
            int selectionIndex)
        {
            return new SourceLineSegment
            {
                ObjectName = name,
                SelectionIndex = selectionIndex,
                IPointName = name + "I",
                JPointName = name + "J",
                StartX = sx,
                StartY = sy,
                StartZ = sz,
                EndX = ex,
                EndY = ey,
                EndZ = ez,
                SectionProperty = "Default",
                StoryName = "Z=0",
                Length = System.Math.Sqrt((ex - sx) * (ex - sx) + (ey - sy) * (ey - sy) + (ez - sz) * (ez - sz))
            };
        }

        private static OffsetPolylineOptions Options()
        {
            return new OffsetPolylineOptions
            {
                CoordinateTolerance = 0.000001,
                PlaneTolerance = 0.000001,
                ZeroLengthTolerance = 0.000001,
                ParallelTolerance = 0.000000001,
                AreaTolerance = 0.000000001,
                MiterLimit = 10
            };
        }

        private static Bounds2D Bounds(IReadOnlyList<OffsetPoint3D> points)
        {
            return new Bounds2D(
                System.Math.Round(points.Min(p => p.X), 6),
                System.Math.Round(points.Max(p => p.X), 6),
                System.Math.Round(points.Min(p => p.Y), 6),
                System.Math.Round(points.Max(p => p.Y), 6));
        }

        private sealed class Bounds2D
        {
            public Bounds2D(double minX, double maxX, double minY, double maxY)
            {
                MinX = minX;
                MaxX = maxX;
                MinY = minY;
                MaxY = maxY;
            }

            public double MinX { get; }
            public double MaxX { get; }
            public double MinY { get; }
            public double MaxY { get; }
        }
    }
}
