using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Media;
using ExcelCSIToolBox.Application.Modelling.PileCaps;
using ExcelCSIToolBox.Core.Contracts.CSI.PileCap;

namespace ExcelCSIToolBoxAddIn.UI.Views
{
    public class PileCapArrangementPreviewControl : FrameworkElement
    {
        public static readonly DependencyProperty GeometryProperty =
            DependencyProperty.Register(
                "Geometry",
                typeof(PileCapGeometry),
                typeof(PileCapArrangementPreviewControl),
                new FrameworkPropertyMetadata(null, FrameworkPropertyMetadataOptions.AffectsRender));

        public static readonly DependencyProperty PileDiameterProperty =
            DependencyProperty.Register(
                "PileDiameter",
                typeof(double),
                typeof(PileCapArrangementPreviewControl),
                new FrameworkPropertyMetadata(800.0, FrameworkPropertyMetadataOptions.AffectsRender));

        public static readonly DependencyProperty BoundaryStrokeProperty =
            DependencyProperty.Register(
                "BoundaryStroke",
                typeof(Brush),
                typeof(PileCapArrangementPreviewControl),
                new FrameworkPropertyMetadata(new SolidColorBrush(Color.FromRgb(190, 48, 48)), FrameworkPropertyMetadataOptions.AffectsRender));

        public static readonly DependencyProperty PileStrokeProperty =
            DependencyProperty.Register(
                "PileStroke",
                typeof(Brush),
                typeof(PileCapArrangementPreviewControl),
                new FrameworkPropertyMetadata(new SolidColorBrush(Color.FromRgb(190, 48, 48)), FrameworkPropertyMetadataOptions.AffectsRender));

        public static readonly DependencyProperty CentrelineStrokeProperty =
            DependencyProperty.Register(
                "CentrelineStroke",
                typeof(Brush),
                typeof(PileCapArrangementPreviewControl),
                new FrameworkPropertyMetadata(new SolidColorBrush(Color.FromRgb(37, 99, 235)), FrameworkPropertyMetadataOptions.AffectsRender));

        public static readonly DependencyProperty IsSelectedProperty =
            DependencyProperty.Register(
                "IsSelected",
                typeof(bool),
                typeof(PileCapArrangementPreviewControl),
                new FrameworkPropertyMetadata(false, FrameworkPropertyMetadataOptions.AffectsRender));

        public PileCapGeometry Geometry
        {
            get { return (PileCapGeometry)GetValue(GeometryProperty); }
            set { SetValue(GeometryProperty, value); }
        }

        public double PileDiameter
        {
            get { return (double)GetValue(PileDiameterProperty); }
            set { SetValue(PileDiameterProperty, value); }
        }

        public Brush BoundaryStroke
        {
            get { return (Brush)GetValue(BoundaryStrokeProperty); }
            set { SetValue(BoundaryStrokeProperty, value); }
        }

        public Brush PileStroke
        {
            get { return (Brush)GetValue(PileStrokeProperty); }
            set { SetValue(PileStrokeProperty, value); }
        }

        public Brush CentrelineStroke
        {
            get { return (Brush)GetValue(CentrelineStrokeProperty); }
            set { SetValue(CentrelineStrokeProperty, value); }
        }

        public bool IsSelected
        {
            get { return (bool)GetValue(IsSelectedProperty); }
            set { SetValue(IsSelectedProperty, value); }
        }

        protected override void OnRender(DrawingContext drawingContext)
        {
            base.OnRender(drawingContext);

            double width = ActualWidth > 1 ? ActualWidth : 128.0;
            double height = ActualHeight > 1 ? ActualHeight : 100.0;
            drawingContext.DrawRectangle(Brushes.White, null, new Rect(0, 0, width, height));

            if (Geometry == null ||
                Geometry.BoundaryVertices == null ||
                Geometry.BoundaryVertices.Count < 3 ||
                Geometry.PileCenters == null ||
                Geometry.PileCenters.Count == 0)
            {
                return;
            }

            IReadOnlyList<PileCapPoint2D> boundary = Rotate(Geometry.BoundaryVertices, Geometry.RotationDegrees);
            IReadOnlyList<PileCapPoint2D> piles = Rotate(Geometry.PileCenters, Geometry.RotationDegrees);
            double pileRadius = Math.Max(1.0, PileDiameter / 2.0);
            ModelBounds bounds = GetBounds(boundary, piles, pileRadius);
            bounds = bounds.Expand(Math.Max(bounds.Width, bounds.Height) * 0.14);

            double scaleX = bounds.Width <= 0 ? 1.0 : width / bounds.Width;
            double scaleY = bounds.Height <= 0 ? 1.0 : height / bounds.Height;
            double scale = Math.Min(scaleX, scaleY);
            if (double.IsNaN(scale) || double.IsInfinity(scale) || scale <= 0)
            {
                scale = 1.0;
            }

            double usedWidth = bounds.Width * scale;
            double usedHeight = bounds.Height * scale;
            double offsetX = (width - usedWidth) / 2.0;
            double offsetY = (height - usedHeight) / 2.0;

            Func<PileCapPoint2D, Point> map = delegate (PileCapPoint2D point)
            {
                double x = offsetX + (point.X - bounds.MinX) * scale;
                double y = offsetY + (bounds.MaxY - point.Y) * scale;
                return new Point(x, y);
            };

            var boundaryPen = new Pen(BoundaryStroke, IsSelected ? 1.8 : 1.6);
            var pilePen = new Pen(PileStroke, 1.35)
            {
                DashStyle = new DashStyle(new[] { 4.0, 3.0 }, 0)
            };
            var centrelinePen = new Pen(CentrelineStroke, 1.35)
            {
                DashStyle = new DashStyle(new[] { 5.0, 3.0 }, 0)
            };

            DrawCentrelines(drawingContext, piles, map, centrelinePen);
            DrawBoundary(drawingContext, boundary, map, boundaryPen);
            DrawPiles(drawingContext, piles, map, pileRadius * scale, pilePen);
        }

        private void DrawCentrelines(
            DrawingContext drawingContext,
            IReadOnlyList<PileCapPoint2D> piles,
            Func<PileCapPoint2D, Point> map,
            Pen pen)
        {
            foreach (LineSegment2D segment in CreateCentrelineSegments(piles))
            {
                drawingContext.DrawLine(pen, map(segment.Start), map(segment.End));
            }
        }

        private IEnumerable<LineSegment2D> CreateCentrelineSegments(IReadOnlyList<PileCapPoint2D> piles)
        {
            if (piles == null)
            {
                yield break;
            }

            switch (Geometry.ArrangementType)
            {
                case PileCapArrangementType.TwoPile:
                    if (piles.Count >= 2)
                    {
                        yield return new LineSegment2D(piles[0], piles[1]);
                    }
                    break;
                case PileCapArrangementType.ThreePile:
                    if (piles.Count >= 3)
                    {
                        yield return new LineSegment2D(piles[0], piles[1]);
                        yield return new LineSegment2D(piles[1], piles[2]);
                        yield return new LineSegment2D(piles[2], piles[0]);
                    }
                    break;
                case PileCapArrangementType.FourPile:
                    if (piles.Count >= 4)
                    {
                        yield return new LineSegment2D(piles[0], piles[1]);
                        yield return new LineSegment2D(piles[1], piles[2]);
                        yield return new LineSegment2D(piles[2], piles[3]);
                        yield return new LineSegment2D(piles[3], piles[0]);
                    }
                    break;
            }
        }

        private static void DrawBoundary(
            DrawingContext drawingContext,
            IReadOnlyList<PileCapPoint2D> boundary,
            Func<PileCapPoint2D, Point> map,
            Pen pen)
        {
            var streamGeometry = new StreamGeometry();
            using (StreamGeometryContext context = streamGeometry.Open())
            {
                context.BeginFigure(map(boundary[0]), false, true);
                for (int i = 1; i < boundary.Count; i++)
                {
                    context.LineTo(map(boundary[i]), true, false);
                }
            }

            streamGeometry.Freeze();
            drawingContext.DrawGeometry(null, pen, streamGeometry);
        }

        private static void DrawPiles(
            DrawingContext drawingContext,
            IReadOnlyList<PileCapPoint2D> piles,
            Func<PileCapPoint2D, Point> map,
            double radius,
            Pen pen)
        {
            double visibleRadius = Math.Max(4.0, radius);
            foreach (PileCapPoint2D pile in piles)
            {
                Point point = map(pile);
                drawingContext.DrawEllipse(null, pen, point, visibleRadius, visibleRadius);
            }
        }

        private static IReadOnlyList<PileCapPoint2D> Rotate(IReadOnlyList<PileCapPoint2D> points, double rotationDegrees)
        {
            var rotated = new List<PileCapPoint2D>();
            if (points == null)
            {
                return rotated;
            }

            double theta = rotationDegrees * Math.PI / 180.0;
            double cos = Math.Cos(theta);
            double sin = Math.Sin(theta);
            foreach (PileCapPoint2D point in points)
            {
                rotated.Add(new PileCapPoint2D(
                    point.X * cos - point.Y * sin,
                    point.X * sin + point.Y * cos));
            }

            return rotated;
        }

        private static ModelBounds GetBounds(
            IReadOnlyList<PileCapPoint2D> boundary,
            IReadOnlyList<PileCapPoint2D> piles,
            double pileRadius)
        {
            var bounds = new ModelBounds();
            foreach (PileCapPoint2D point in boundary)
            {
                bounds.Include(point.X, point.Y);
            }

            foreach (PileCapPoint2D point in piles)
            {
                bounds.Include(point.X - pileRadius, point.Y - pileRadius);
                bounds.Include(point.X + pileRadius, point.Y + pileRadius);
            }

            return bounds;
        }

        private struct ModelBounds
        {
            private bool _initialized;

            public double MinX { get; private set; }

            public double MaxX { get; private set; }

            public double MinY { get; private set; }

            public double MaxY { get; private set; }

            public double Width
            {
                get { return Math.Max(1.0, MaxX - MinX); }
            }

            public double Height
            {
                get { return Math.Max(1.0, MaxY - MinY); }
            }

            public void Include(double x, double y)
            {
                if (!_initialized)
                {
                    MinX = MaxX = x;
                    MinY = MaxY = y;
                    _initialized = true;
                    return;
                }

                MinX = Math.Min(MinX, x);
                MaxX = Math.Max(MaxX, x);
                MinY = Math.Min(MinY, y);
                MaxY = Math.Max(MaxY, y);
            }

            public ModelBounds Expand(double amount)
            {
                return new ModelBounds
                {
                    _initialized = true,
                    MinX = MinX - amount,
                    MaxX = MaxX + amount,
                    MinY = MinY - amount,
                    MaxY = MaxY + amount
                };
            }
        }
    }
}
