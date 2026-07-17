using System;
using System.Collections.Generic;
using System.Globalization;
using System.Windows;
using System.Windows.Media;
using ExcelCSIToolBox.Application.Modelling.PileCaps;
using ExcelCSIToolBox.Core.Contracts.CSI.PileCap;

namespace ExcelCSIToolBoxAddIn.UI.Views
{
    public class PileCapPreviewCanvas : FrameworkElement
    {
        public static readonly DependencyProperty GeometryProperty =
            DependencyProperty.Register(
                "Geometry",
                typeof(PileCapGeometry),
                typeof(PileCapPreviewCanvas),
                new FrameworkPropertyMetadata(null, FrameworkPropertyMetadataOptions.AffectsRender));

        public static readonly DependencyProperty PileDiameterProperty =
            DependencyProperty.Register(
                "PileDiameter",
                typeof(double),
                typeof(PileCapPreviewCanvas),
                new FrameworkPropertyMetadata(800.0, FrameworkPropertyMetadataOptions.AffectsRender));

        public static readonly DependencyProperty PileCapThicknessProperty =
            DependencyProperty.Register(
                "PileCapThickness",
                typeof(double),
                typeof(PileCapPreviewCanvas),
                new FrameworkPropertyMetadata(1500.0, FrameworkPropertyMetadataOptions.AffectsRender));

        public static readonly DependencyProperty EdgeDistanceProperty =
            DependencyProperty.Register(
                "EdgeDistance",
                typeof(double),
                typeof(PileCapPreviewCanvas),
                new FrameworkPropertyMetadata(150.0, FrameworkPropertyMetadataOptions.AffectsRender));

        public static readonly DependencyProperty ShowMeshProperty =
            DependencyProperty.Register(
                "ShowMesh",
                typeof(bool),
                typeof(PileCapPreviewCanvas),
                new FrameworkPropertyMetadata(false, FrameworkPropertyMetadataOptions.AffectsRender));

        public static readonly DependencyProperty ShowAxesProperty =
            DependencyProperty.Register(
                "ShowAxes",
                typeof(bool),
                typeof(PileCapPreviewCanvas),
                new FrameworkPropertyMetadata(false, FrameworkPropertyMetadataOptions.AffectsRender));

        public static readonly DependencyProperty ShowDimensionsProperty =
            DependencyProperty.Register(
                "ShowDimensions",
                typeof(bool),
                typeof(PileCapPreviewCanvas),
                new FrameworkPropertyMetadata(true, FrameworkPropertyMetadataOptions.AffectsRender));

        public static readonly DependencyProperty ShowConstructionGeometryProperty =
            DependencyProperty.Register(
                "ShowConstructionGeometry",
                typeof(bool),
                typeof(PileCapPreviewCanvas),
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

        public double PileCapThickness
        {
            get { return (double)GetValue(PileCapThicknessProperty); }
            set { SetValue(PileCapThicknessProperty, value); }
        }

        public double EdgeDistance
        {
            get { return (double)GetValue(EdgeDistanceProperty); }
            set { SetValue(EdgeDistanceProperty, value); }
        }

        public bool ShowMesh
        {
            get { return (bool)GetValue(ShowMeshProperty); }
            set { SetValue(ShowMeshProperty, value); }
        }

        public bool ShowAxes
        {
            get { return (bool)GetValue(ShowAxesProperty); }
            set { SetValue(ShowAxesProperty, value); }
        }

        public bool ShowDimensions
        {
            get { return (bool)GetValue(ShowDimensionsProperty); }
            set { SetValue(ShowDimensionsProperty, value); }
        }

        public bool ShowConstructionGeometry
        {
            get { return (bool)GetValue(ShowConstructionGeometryProperty); }
            set { SetValue(ShowConstructionGeometryProperty, value); }
        }

        protected override void OnRender(DrawingContext drawingContext)
        {
            base.OnRender(drawingContext);

            double canvasWidth = ActualWidth > 20 ? ActualWidth : 520;
            double canvasHeight = ActualHeight > 20 ? ActualHeight : 360;
            drawingContext.DrawRectangle(Brushes.White, null, new Rect(0, 0, canvasWidth, canvasHeight));

            if (Geometry == null || Geometry.BoundaryVertices == null || Geometry.BoundaryVertices.Count < 3)
            {
                DrawText(drawingContext, "Preview", new Point(12, 12), Brushes.Gray, 11);
                return;
            }

            IReadOnlyList<PileCapPoint2D> boundary = Rotate(Geometry.BoundaryVertices, Geometry.RotationDegrees);
            IReadOnlyList<PileCapPoint2D> piles = Rotate(Geometry.PileCenters, Geometry.RotationDegrees);
            PileCapPoint2D selectedPoint = Rotate(new[] { Geometry.SelectedPoint }, Geometry.RotationDegrees)[0];

            double pileRadiusModel = Math.Max(1.0, PileDiameter / 2.0);
            Rect geometryBounds = GetGeometryBounds(boundary, piles, pileRadiusModel);
            Rect displayBounds = ExpandBounds(
                geometryBounds,
                ShowDimensions ? Math.Max(160.0, Math.Max(geometryBounds.Width, geometryBounds.Height) * 0.18) : Math.Max(40.0, pileRadiusModel));

            double padding = ShowDimensions ? 18.0 : 6.0;
            double scaleX = displayBounds.Width <= 0 ? 1 : (canvasWidth - padding * 2.0) / displayBounds.Width;
            double scaleY = displayBounds.Height <= 0 ? 1 : (canvasHeight - padding * 2.0) / displayBounds.Height;
            double scale = Math.Min(scaleX, scaleY);
            if (double.IsNaN(scale) || double.IsInfinity(scale) || scale <= 0)
            {
                scale = 1;
            }

            Func<PileCapPoint2D, Point> map = delegate (PileCapPoint2D point)
            {
                double x = padding + (point.X - displayBounds.Left) * scale;
                double y = canvasHeight - padding - (point.Y - displayBounds.Top) * scale;
                return new Point(x, y);
            };

            Brush capFill = new SolidColorBrush(Color.FromRgb(247, 248, 250));
            Pen capPen = new Pen(new SolidColorBrush(Color.FromRgb(55, 65, 81)), ShowDimensions ? 1.6 : 1.2);
            Brush pileBrush = new SolidColorBrush(Color.FromRgb(31, 41, 55));
            Brush selectedPointBrush = new SolidColorBrush(Color.FromRgb(37, 99, 235));

            if (ShowConstructionGeometry)
            {
                DrawTemporaryMonoCaps(drawingContext, map);
            }

            DrawPolygon(drawingContext, boundary, map, capFill, capPen);

            if (ShowMesh)
            {
                DrawMesh(drawingContext, map);
            }

            double pileRadius = Math.Max(3.0, Math.Min(18.0, PileDiameter * scale / 2.0));
            foreach (PileCapPoint2D pile in piles)
            {
                Point point = map(pile);
                drawingContext.DrawEllipse(pileBrush, new Pen(Brushes.White, 0.8), point, pileRadius, pileRadius);
            }

            Point center = map(selectedPoint);
            double centerHalfSize = ShowDimensions ? 5.0 : 3.2;
            drawingContext.DrawRectangle(
                selectedPointBrush,
                new Pen(Brushes.White, 0.8),
                new Rect(center.X - centerHalfSize, center.Y - centerHalfSize, centerHalfSize * 2.0, centerHalfSize * 2.0));

            if (ShowAxes)
            {
                DrawLocalXAxis(drawingContext, geometryBounds, map);
            }

            if (ShowDimensions)
            {
                DrawDimensionLines(drawingContext, geometryBounds, piles, map);
                DrawDimensionReadout(drawingContext);
            }
        }

        private void DrawTemporaryMonoCaps(DrawingContext drawingContext, Func<PileCapPoint2D, Point> map)
        {
            if (Geometry.TemporaryMonoCaps == null || Geometry.TemporaryMonoCaps.Count == 0)
            {
                return;
            }

            var brush = new SolidColorBrush(Color.FromRgb(156, 163, 175));
            var pen = new Pen(brush, 0.8) { DashStyle = DashStyles.Dash };
            foreach (Rectangle2D temporaryMonoCap in Geometry.TemporaryMonoCaps)
            {
                DrawPolyline(
                    drawingContext,
                    Rotate(temporaryMonoCap.GetCornersClockwise(), Geometry.RotationDegrees),
                    map,
                    pen,
                    true);
            }
        }

        private void DrawMesh(DrawingContext drawingContext, Func<PileCapPoint2D, Point> map)
        {
            if (Geometry.MeshAreas == null)
            {
                return;
            }

            var meshPen = new Pen(new SolidColorBrush(Color.FromRgb(203, 213, 225)), 0.7);
            foreach (PileCapMeshArea meshArea in Geometry.MeshAreas)
            {
                DrawPolyline(drawingContext, Rotate(meshArea.Points, Geometry.RotationDegrees), map, meshPen, true);
            }
        }

        private void DrawDimensionLines(
            DrawingContext drawingContext,
            Rect geometryBounds,
            IReadOnlyList<PileCapPoint2D> piles,
            Func<PileCapPoint2D, Point> map)
        {
            double extension = Math.Max(120.0, Math.Max(geometryBounds.Width, geometryBounds.Height) * 0.08);
            var pen = new Pen(new SolidColorBrush(Color.FromRgb(107, 114, 128)), 0.9);
            Brush brush = new SolidColorBrush(Color.FromRgb(75, 85, 99));

            PileCapPoint2D bottomLeft = new PileCapPoint2D(geometryBounds.Left, geometryBounds.Top - extension);
            PileCapPoint2D bottomRight = new PileCapPoint2D(geometryBounds.Right, geometryBounds.Top - extension);
            DrawScreenDimension(
                drawingContext,
                map(bottomLeft),
                map(bottomRight),
                "X " + Format(Geometry.CapWidthXMillimeters) + " mm",
                brush,
                pen);

            PileCapPoint2D rightBottom = new PileCapPoint2D(geometryBounds.Right + extension, geometryBounds.Top);
            PileCapPoint2D rightTop = new PileCapPoint2D(geometryBounds.Right + extension, geometryBounds.Bottom);
            DrawScreenDimension(
                drawingContext,
                map(rightBottom),
                map(rightTop),
                "Y " + Format(Geometry.CapLengthYMillimeters) + " mm",
                brush,
                pen);

            if (piles != null && piles.Count >= 2 && Geometry.SpacingXMillimeters > 0)
            {
                DrawScreenDimension(
                    drawingContext,
                    map(piles[0]),
                    map(piles[1]),
                    Geometry.ArrangementType == PileCapArrangementType.ThreePile
                        ? "S " + Format(Geometry.SpacingXMillimeters) + " mm"
                        : "Sx " + Format(Geometry.SpacingXMillimeters) + " mm",
                    brush,
                    pen);
            }
        }

        private void DrawScreenDimension(
            DrawingContext drawingContext,
            Point start,
            Point end,
            string label,
            Brush brush,
            Pen pen)
        {
            drawingContext.DrawLine(pen, start, end);

            Vector direction = end - start;
            if (direction.Length <= 0)
            {
                return;
            }

            direction.Normalize();
            Vector tick = new Vector(-direction.Y, direction.X) * 5.0;
            drawingContext.DrawLine(pen, start - tick, start + tick);
            drawingContext.DrawLine(pen, end - tick, end + tick);

            Point midpoint = new Point((start.X + end.X) / 2.0, (start.Y + end.Y) / 2.0);
            DrawText(drawingContext, label, midpoint + new Vector(5, -15), brush, 10);
        }

        private void DrawDimensionReadout(DrawingContext drawingContext)
        {
            var lines = new List<string>
            {
                "D " + Format(PileDiameter) + " mm",
                "X " + Format(Geometry.CapWidthXMillimeters) + " mm",
                "Y " + Format(Geometry.CapLengthYMillimeters) + " mm"
            };

            if (Geometry.ArrangementType == PileCapArrangementType.ThreePile)
            {
                lines.Add("S " + Format(Geometry.SpacingXMillimeters) + " mm");
            }
            else
            {
                if (Geometry.SpacingXMillimeters > 0)
                {
                    lines.Add("Sx " + Format(Geometry.SpacingXMillimeters) + " mm");
                }

                if (Geometry.SpacingYMillimeters > 0)
                {
                    lines.Add("Sy " + Format(Geometry.SpacingYMillimeters) + " mm");
                }
            }

            lines.Add("E " + Format(EdgeDistance) + " mm");
            lines.Add("t " + Format(PileCapThickness) + " mm");
            lines.Add("Rotation " + Format(Geometry.RotationDegrees) + " deg");

            Brush brush = new SolidColorBrush(Color.FromRgb(75, 85, 99));
            for (int i = 0; i < lines.Count; i++)
            {
                DrawText(drawingContext, lines[i], new Point(12, 10 + i * 15), brush, 10);
            }
        }

        private void DrawLocalXAxis(DrawingContext drawingContext, Rect geometryBounds, Func<PileCapPoint2D, Point> map)
        {
            double y = geometryBounds.Bottom + Math.Max(100.0, geometryBounds.Height * 0.08);
            double x0 = geometryBounds.Left;
            double x1 = geometryBounds.Left + Math.Max(220.0, geometryBounds.Width * 0.18);
            var pen = new Pen(new SolidColorBrush(Color.FromRgb(31, 122, 58)), 1.2);
            DrawArrow(drawingContext, map(new PileCapPoint2D(x0, y)), map(new PileCapPoint2D(x1, y)), pen);
            DrawText(drawingContext, "X", map(new PileCapPoint2D(x1, y)) + new Vector(4, -8), pen.Brush, 10);
        }

        private static void DrawPolygon(
            DrawingContext drawingContext,
            IReadOnlyList<PileCapPoint2D> points,
            Func<PileCapPoint2D, Point> map,
            Brush fill,
            Pen stroke)
        {
            if (points == null || points.Count < 3)
            {
                return;
            }

            var geometry = new StreamGeometry();
            using (StreamGeometryContext context = geometry.Open())
            {
                context.BeginFigure(map(points[0]), true, true);
                for (int i = 1; i < points.Count; i++)
                {
                    context.LineTo(map(points[i]), true, false);
                }
            }

            geometry.Freeze();
            drawingContext.DrawGeometry(fill, stroke, geometry);
        }

        private static void DrawPolyline(
            DrawingContext drawingContext,
            IReadOnlyList<PileCapPoint2D> points,
            Func<PileCapPoint2D, Point> map,
            Pen stroke,
            bool close)
        {
            if (points == null || points.Count < 2)
            {
                return;
            }

            var geometry = new StreamGeometry();
            using (StreamGeometryContext context = geometry.Open())
            {
                context.BeginFigure(map(points[0]), false, close);
                for (int i = 1; i < points.Count; i++)
                {
                    context.LineTo(map(points[i]), true, false);
                }
            }

            geometry.Freeze();
            drawingContext.DrawGeometry(null, stroke, geometry);
        }

        private static void DrawArrow(DrawingContext drawingContext, Point start, Point end, Pen pen)
        {
            drawingContext.DrawLine(pen, start, end);
            Vector direction = start - end;
            if (direction.Length <= 0)
            {
                return;
            }

            direction.Normalize();
            Vector normal = new Vector(-direction.Y, direction.X);
            Point head1 = end + direction * 7.0 + normal * 3.5;
            Point head2 = end + direction * 7.0 - normal * 3.5;
            drawingContext.DrawLine(pen, end, head1);
            drawingContext.DrawLine(pen, end, head2);
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

        private static Rect GetGeometryBounds(
            IReadOnlyList<PileCapPoint2D> boundary,
            IReadOnlyList<PileCapPoint2D> piles,
            double pileRadius)
        {
            bool initialized = false;
            double minX = 0;
            double maxX = 0;
            double minY = 0;
            double maxY = 0;
            Action<double, double> addCoordinate = delegate (double x, double y)
            {
                if (!initialized)
                {
                    minX = maxX = x;
                    minY = maxY = y;
                    initialized = true;
                    return;
                }

                minX = Math.Min(minX, x);
                maxX = Math.Max(maxX, x);
                minY = Math.Min(minY, y);
                maxY = Math.Max(maxY, y);
            };

            if (boundary != null)
            {
                foreach (PileCapPoint2D point in boundary)
                {
                    addCoordinate(point.X, point.Y);
                }
            }

            if (piles != null)
            {
                foreach (PileCapPoint2D point in piles)
                {
                    addCoordinate(point.X - pileRadius, point.Y - pileRadius);
                    addCoordinate(point.X + pileRadius, point.Y + pileRadius);
                }
            }

            return initialized
                ? new Rect(minX, minY, Math.Max(1.0, maxX - minX), Math.Max(1.0, maxY - minY))
                : new Rect(-1, -1, 2, 2);
        }

        private static Rect ExpandBounds(Rect bounds, double amount)
        {
            return new Rect(
                bounds.Left - amount,
                bounds.Top - amount,
                bounds.Width + amount * 2.0,
                bounds.Height + amount * 2.0);
        }

        private static void DrawText(DrawingContext drawingContext, string text, Point point, Brush brush, double fontSize)
        {
            var formattedText = new FormattedText(
                text ?? string.Empty,
                CultureInfo.CurrentCulture,
                FlowDirection.LeftToRight,
                new Typeface("Segoe UI"),
                fontSize,
                brush,
                1.0);
            drawingContext.DrawText(formattedText, point);
        }

        private static string Format(double value)
        {
            return Math.Round(value, 0, MidpointRounding.AwayFromZero).ToString("0", CultureInfo.InvariantCulture);
        }
    }
}
