using System.Collections.Generic;

namespace ExcelCSIToolBox.Application.Modelling.OffsetPolylines
{
    public struct OffsetPoint3D
    {
        public OffsetPoint3D(double x, double y, double z)
        {
            X = x;
            Y = y;
            Z = z;
        }

        public double X { get; }
        public double Y { get; }
        public double Z { get; }
    }

    public struct OffsetPoint2D
    {
        public OffsetPoint2D(double u, double v)
        {
            U = u;
            V = v;
        }

        public double U { get; }
        public double V { get; }
    }

    public sealed class OffsetPolylineOptions
    {
        public OffsetPolylineOptions()
        {
            CoordinateTolerance = 0.000001;
            PlaneTolerance = 0.000001;
            ZeroLengthTolerance = 0.000001;
            ParallelTolerance = 0.000000001;
            AreaTolerance = 0.000000001;
            MiterLimit = 10.0;
        }

        public double CoordinateTolerance { get; set; }
        public double PlaneTolerance { get; set; }
        public double ZeroLengthTolerance { get; set; }
        public double ParallelTolerance { get; set; }
        public double AreaTolerance { get; set; }
        public double MiterLimit { get; set; }
    }

    public sealed class SourceLineSegment
    {
        public string ObjectName { get; set; }
        public int SelectionIndex { get; set; }
        public string IPointName { get; set; }
        public string JPointName { get; set; }
        public double StartX { get; set; }
        public double StartY { get; set; }
        public double StartZ { get; set; }
        public double EndX { get; set; }
        public double EndY { get; set; }
        public double EndZ { get; set; }
        public string SectionProperty { get; set; }
        public string StoryName { get; set; }
        public double Length { get; set; }

        public OffsetPoint3D StartPoint
        {
            get { return new OffsetPoint3D(StartX, StartY, StartZ); }
        }

        public OffsetPoint3D EndPoint
        {
            get { return new OffsetPoint3D(EndX, EndY, EndZ); }
        }
    }

    public sealed class OrderedLineSegment
    {
        public string SourceObjectName { get; set; }
        public int SourceSelectionIndex { get; set; }
        public int OrderedIndex { get; set; }
        public OffsetPoint3D OriginalStartPoint { get; set; }
        public OffsetPoint3D OriginalEndPoint { get; set; }
        public OffsetPoint3D OrderedStartPoint { get; set; }
        public OffsetPoint3D OrderedEndPoint { get; set; }
        public bool IsReversedDuringOrdering { get; set; }
        public string SourceSectionProperty { get; set; }
    }

    public sealed class OffsetLineSegment
    {
        public string SourceObjectName { get; set; }
        public int SourceOrderedIndex { get; set; }
        public int ResultIndex { get; set; }
        public string NewObjectName { get; set; }
        public double StartX { get; set; }
        public double StartY { get; set; }
        public double StartZ { get; set; }
        public double EndX { get; set; }
        public double EndY { get; set; }
        public double EndZ { get; set; }
        public double OffsetDistance { get; set; }
        public string ResultType { get; set; }
        public string SourceSectionProperty { get; set; }
        public bool IsReversedDuringOrdering { get; set; }

        public OffsetPoint3D StartPoint
        {
            get { return new OffsetPoint3D(StartX, StartY, StartZ); }
        }

        public OffsetPoint3D EndPoint
        {
            get { return new OffsetPoint3D(EndX, EndY, EndZ); }
        }
    }

    public sealed class OffsetPolylineResult
    {
        public bool IsValid { get; set; }
        public string ValidationMessage { get; set; }
        public double OffsetDistance { get; set; }
        public string OffsetDirection { get; set; }
        public string ResultType { get; set; }
        public string PolygonOrientation { get; set; }
        public OffsetPoint3D PlaneOrigin { get; set; }
        public OffsetPoint3D PlaneNormal { get; set; }
        public OffsetPoint3D PlaneXAxis { get; set; }
        public OffsetPoint3D PlaneYAxis { get; set; }
        public IReadOnlyList<SourceLineSegment> OriginalSegments { get; set; }
        public IReadOnlyList<OrderedLineSegment> OrderedSegments { get; set; }
        public IReadOnlyList<OffsetLineSegment> ResultSegments { get; set; }
        public IReadOnlyList<OffsetPoint3D> OriginalVertices { get; set; }
        public IReadOnlyList<OffsetPoint3D> ResultVertices { get; set; }
        public double SourceArea { get; set; }
        public double ResultArea { get; set; }
        public string GroupName { get; set; }
        public int DetectedVertexCount { get; set; }
    }
}
