using System.Collections.Generic;
using ExcelCSIToolBox.Core.Contracts.CSI.PileCap;

namespace ExcelCSIToolBox.Application.Modelling.PileCaps
{
    public class PileCapInputParameters
    {
        public PileCapArrangementType ArrangementType { get; set; }

        public double PileDiameterMillimeters { get; set; }

        public double PileLengthMillimeters { get; set; }

        public double RotationDegrees { get; set; }

        public bool AutoSpacing { get; set; }

        public double PileSpacingMillimeters { get; set; }

        public double SpacingXMillimeters { get; set; }

        public double SpacingYMillimeters { get; set; }

        public double PileCapThicknessMillimeters { get; set; }

        public double EdgeDistanceMillimeters { get; set; }

        public string PileMaterial { get; set; }

        public string PileCapMaterial { get; set; }
    }

    public class PileCapGeometry
    {
        public PileCapGeometry(
            PileCapArrangementType arrangementType,
            IReadOnlyList<PileCapPoint2D> pileCenters,
            IReadOnlyList<PileCapPoint2D> boundaryVertices,
            IReadOnlyList<PileCapMeshArea> meshAreas,
            double capWidthXMillimeters,
            double capLengthYMillimeters,
            double spacingXMillimeters,
            double spacingYMillimeters,
            double rotationDegrees,
            PileCapPoint2D selectedPoint,
            IReadOnlyList<Rectangle2D> temporaryMonoCaps,
            IReadOnlyList<LineSegment2D> boundarySegments)
        {
            ArrangementType = arrangementType;
            PileCenters = pileCenters;
            BoundaryVertices = boundaryVertices;
            MeshAreas = meshAreas;
            CapWidthXMillimeters = capWidthXMillimeters;
            CapLengthYMillimeters = capLengthYMillimeters;
            SpacingXMillimeters = spacingXMillimeters;
            SpacingYMillimeters = spacingYMillimeters;
            RotationDegrees = rotationDegrees;
            SelectedPoint = selectedPoint;
            TemporaryMonoCaps = temporaryMonoCaps ?? new List<Rectangle2D>();
            BoundarySegments = boundarySegments ?? CreateBoundarySegments(boundaryVertices);
        }

        public PileCapGeometry(
            PileCapArrangementType arrangementType,
            IReadOnlyList<PileCapPoint2D> pileCenters,
            IReadOnlyList<PileCapPoint2D> boundaryVertices,
            IReadOnlyList<PileCapMeshArea> meshAreas,
            double capWidthXMillimeters,
            double capLengthYMillimeters,
            double spacingXMillimeters,
            double spacingYMillimeters,
            double rotationDegrees)
            : this(
                arrangementType,
                pileCenters,
                boundaryVertices,
                meshAreas,
                capWidthXMillimeters,
                capLengthYMillimeters,
                spacingXMillimeters,
                spacingYMillimeters,
                rotationDegrees,
                new PileCapPoint2D(0, 0),
                null,
                null)
        {
        }

        public PileCapArrangementType ArrangementType { get; private set; }

        public IReadOnlyList<PileCapPoint2D> PileCenters { get; private set; }

        public IReadOnlyList<PileCapPoint2D> BoundaryVertices { get; private set; }

        public IReadOnlyList<PileCapMeshArea> MeshAreas { get; private set; }

        public double CapWidthXMillimeters { get; private set; }

        public double CapLengthYMillimeters { get; private set; }

        public double SpacingXMillimeters { get; private set; }

        public double SpacingYMillimeters { get; private set; }

        public double RotationDegrees { get; private set; }

        public PileCapPoint2D SelectedPoint { get; private set; }

        public IReadOnlyList<Rectangle2D> TemporaryMonoCaps { get; private set; }

        public IReadOnlyList<LineSegment2D> BoundarySegments { get; private set; }

        private static IReadOnlyList<LineSegment2D> CreateBoundarySegments(IReadOnlyList<PileCapPoint2D> vertices)
        {
            var segments = new List<LineSegment2D>();
            if (vertices == null || vertices.Count < 2)
            {
                return segments;
            }

            for (int i = 0; i < vertices.Count; i++)
            {
                PileCapPoint2D start = vertices[i];
                PileCapPoint2D end = vertices[(i + 1) % vertices.Count];
                segments.Add(new LineSegment2D(start, end));
            }

            return segments;
        }
    }

    public class PileCapMeshArea
    {
        public PileCapMeshArea(IReadOnlyList<PileCapPoint2D> points)
        {
            Points = points;
        }

        public IReadOnlyList<PileCapPoint2D> Points { get; private set; }
    }

    public struct PileCapPoint2D
    {
        public PileCapPoint2D(double x, double y)
        {
            X = x;
            Y = y;
        }

        public double X { get; private set; }

        public double Y { get; private set; }
    }

    public struct PileCapPoint3D
    {
        public PileCapPoint3D(double x, double y, double z)
        {
            X = x;
            Y = y;
            Z = z;
        }

        public double X { get; private set; }

        public double Y { get; private set; }

        public double Z { get; private set; }
    }

    public sealed class Rectangle2D
    {
        public Rectangle2D(double left, double right, double bottom, double top)
        {
            Left = left;
            Right = right;
            Bottom = bottom;
            Top = top;
        }

        public double Left { get; private set; }

        public double Right { get; private set; }

        public double Bottom { get; private set; }

        public double Top { get; private set; }

        public double Width
        {
            get { return Right - Left; }
        }

        public double Height
        {
            get { return Top - Bottom; }
        }

        public IReadOnlyList<PileCapPoint2D> GetCornersClockwise()
        {
            return new[]
            {
                new PileCapPoint2D(Left, Bottom),
                new PileCapPoint2D(Right, Bottom),
                new PileCapPoint2D(Right, Top),
                new PileCapPoint2D(Left, Top)
            };
        }
    }

    public sealed class LineSegment2D
    {
        public LineSegment2D(PileCapPoint2D start, PileCapPoint2D end)
        {
            Start = start;
            End = end;
        }

        public PileCapPoint2D Start { get; private set; }

        public PileCapPoint2D End { get; private set; }
    }

    public sealed class ThreePileCapGeometry
    {
        public ThreePileCapGeometry(
            IReadOnlyList<PileCapPoint2D> pileCenters,
            PileCapPoint2D selectedPoint,
            IReadOnlyList<Rectangle2D> temporaryMonoCaps,
            IReadOnlyList<PileCapPoint2D> finalBoundaryVertices,
            IReadOnlyList<LineSegment2D> finalBoundarySegments)
        {
            PileCenters = pileCenters;
            SelectedPoint = selectedPoint;
            TemporaryMonoCaps = temporaryMonoCaps;
            FinalBoundaryVertices = finalBoundaryVertices;
            FinalBoundarySegments = finalBoundarySegments;
        }

        public IReadOnlyList<PileCapPoint2D> PileCenters { get; private set; }

        public PileCapPoint2D SelectedPoint { get; private set; }

        public IReadOnlyList<Rectangle2D> TemporaryMonoCaps { get; private set; }

        public IReadOnlyList<PileCapPoint2D> FinalBoundaryVertices { get; private set; }

        public IReadOnlyList<LineSegment2D> FinalBoundarySegments { get; private set; }
    }
}
