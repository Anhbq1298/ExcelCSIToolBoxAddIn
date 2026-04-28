using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Effects;
using System.Windows.Shapes;
using ExcelCSIToolBoxAddIn.Data.DTOs;

namespace ExcelCSIToolBoxAddIn.UI.Helpers
{
    /// <summary>
    /// Renders structural cross-section shapes onto a WPF Canvas.
    /// Canvas is assumed to be 100×100 logical units, stretched via Viewbox.
    /// </summary>
    internal static class SectionShapeRenderer
    {
        private static readonly Brush StrokeBrush = new SolidColorBrush(Color.FromRgb(30, 50, 90));
        private const double StrokeThickness = 1.5;

        private static Brush MakeFill()
        {
            var grad = new LinearGradientBrush();
            grad.StartPoint = new Point(0, 0);
            grad.EndPoint = new Point(1, 1);
            grad.GradientStops.Add(new GradientStop(Color.FromArgb(220, 100, 149, 237), 0));
            grad.GradientStops.Add(new GradientStop(Color.FromArgb(180, 50, 90, 180), 1));
            grad.Freeze();
            return grad;
        }

        private static DropShadowEffect MakeShadow() => new DropShadowEffect
        {
            Color = Colors.SteelBlue,
            BlurRadius = 6,
            ShadowDepth = 2,
            Opacity = 0.4
        };

        public static void Render(Canvas canvas, CSISapModelFrameSectionDetailDTO dto)
        {
            canvas.Children.Clear();
            DrawPreviewGrid(canvas);

            if (dto == null)
            {
                AddLabel(canvas, "No section selected");
                return;
            }

            double cx = 50, cy = 50;
            var dims = dto.Dimensions;

            double t3 = Get(dims, 100, "Total depth ( t3 )", "Depth ( t3 )", "Outside diameter ( t3 )", "Diameter ( t3 )");
            double scale = t3 > 0 ? 80.0 / t3 : 1.0;

            Path shapePath;

            switch (dto.ShapeType)
            {
                case FrameSectionShapeType.Pipe:
                    shapePath = BuildPipe(cx, cy, t3 * scale, Get(dims, "Wall thickness ( tw )") * scale);
                    break;
                case FrameSectionShapeType.Circular:
                    shapePath = BuildCircle(cx, cy, t3 * scale);
                    break;
                case FrameSectionShapeType.Rectangular:
                    shapePath = BuildRect(cx, cy, Rescale(t3, Get(dims, "Width ( t2 )", t3), ref scale, 80));
                    break;
                case FrameSectionShapeType.Tube:
                    shapePath = BuildTube(cx, cy, t3, Get(dims, "Flange width ( t2 )", t3), Get(dims, "Flange thickness ( tf )"), Get(dims, "Web thickness ( tw )"), ref scale);
                    break;
                case FrameSectionShapeType.I:
                    shapePath = BuildISection(cx, cy, t3, Get(dims, "Top flange width ( t2 )", t3), Get(dims, "Top flange thickness ( tf )"), Get(dims, "Web thickness ( tw )"), Get(dims, "Bottom flange width ( t2b )", Get(dims, "Top flange width ( t2 )", t3)), Get(dims, "Bottom flange thickness ( tfb )", Get(dims, "Top flange thickness ( tf )")), ref scale);
                    break;
                case FrameSectionShapeType.Channel:
                    shapePath = BuildChannel(cx, cy, t3, Get(dims, "Flange width ( t2 )", t3), Get(dims, "Flange thickness ( tf )"), Get(dims, "Web thickness ( tw )"), ref scale);
                    break;
                case FrameSectionShapeType.Angle:
                    shapePath = BuildAngle(cx, cy, t3, Get(dims, "Flange width ( t2 )", t3), Get(dims, "Flange thickness ( tf )"), Get(dims, "Web thickness ( tw )"), ref scale);
                    break;
                default:
                    AddLabel(canvas, $"{dto.ShapeType}\n(no preview)");
                    return;
            }

            shapePath.Stroke = StrokeBrush;
            shapePath.StrokeThickness = StrokeThickness;
            shapePath.Fill = MakeFill();
            shapePath.Effect = MakeShadow();
            canvas.Children.Add(shapePath);
        }

        // --- Shape builders ---

        private static Path BuildPipe(double cx, double cy, double d, double tw)
        {
            var group = new GeometryGroup { FillRule = FillRule.EvenOdd };
            group.Children.Add(new EllipseGeometry(new Point(cx, cy), d / 2, d / 2));
            double inner = d / 2 - tw;
            if (tw > 0 && inner > 0)
                group.Children.Add(new EllipseGeometry(new Point(cx, cy), inner, inner));
            return new Path { Data = group };
        }

        private static Path BuildCircle(double cx, double cy, double d)
        {
            return new Path { Data = new EllipseGeometry(new Point(cx, cy), d / 2, d / 2) };
        }

        private static Path BuildRect(double cx, double cy, (double d, double b) dims)
        {
            return new Path { Data = new RectangleGeometry(new Rect(cx - dims.b / 2, cy - dims.d / 2, dims.b, dims.d)) };
        }

        private static Path BuildTube(double cx, double cy, double t3Raw, double t2Raw, double tfRaw, double twRaw, ref double scale)
        {
            LimitScale(ref scale, t3Raw, t2Raw, 80);
            double d = t3Raw * scale, b = t2Raw * scale, tf = tfRaw * scale, tw = twRaw * scale;
            var group = new GeometryGroup { FillRule = FillRule.EvenOdd };
            group.Children.Add(new RectangleGeometry(new Rect(cx - b / 2, cy - d / 2, b, d)));
            if (b - 2 * tw > 0 && d - 2 * tf > 0)
                group.Children.Add(new RectangleGeometry(new Rect(cx - b / 2 + tw, cy - d / 2 + tf, b - 2 * tw, d - 2 * tf)));
            return new Path { Data = group };
        }

        private static Path BuildISection(double cx, double cy, double t3Raw, double t2Raw, double tfRaw, double twRaw, double t2bRaw, double tfbRaw, ref double scale)
        {
            LimitScale(ref scale, t3Raw, Math.Max(t2Raw, t2bRaw), 80);
            double d = t3Raw * scale, t2 = t2Raw * scale, tf = tfRaw * scale, tw = twRaw * scale, t2b = t2bRaw * scale, tfb = tfbRaw * scale;

            var pf = new PathFigure { IsClosed = true, StartPoint = new Point(cx - t2 / 2, cy - d / 2) };
            pf.Segments.Add(Seg(cx + t2 / 2, cy - d / 2));
            pf.Segments.Add(Seg(cx + t2 / 2, cy - d / 2 + tf));
            pf.Segments.Add(Seg(cx + tw / 2, cy - d / 2 + tf));
            pf.Segments.Add(Seg(cx + tw / 2, cy + d / 2 - tfb));
            pf.Segments.Add(Seg(cx + t2b / 2, cy + d / 2 - tfb));
            pf.Segments.Add(Seg(cx + t2b / 2, cy + d / 2));
            pf.Segments.Add(Seg(cx - t2b / 2, cy + d / 2));
            pf.Segments.Add(Seg(cx - t2b / 2, cy + d / 2 - tfb));
            pf.Segments.Add(Seg(cx - tw / 2, cy + d / 2 - tfb));
            pf.Segments.Add(Seg(cx - tw / 2, cy - d / 2 + tf));
            pf.Segments.Add(Seg(cx - t2 / 2, cy - d / 2 + tf));

            return new Path { Data = new PathGeometry(new[] { pf }) };
        }

        private static Path BuildChannel(double cx, double cy, double t3Raw, double t2Raw, double tfRaw, double twRaw, ref double scale)
        {
            LimitScale(ref scale, t3Raw, t2Raw, 80);
            double d = t3Raw * scale, b = t2Raw * scale, tf = tfRaw * scale, tw = twRaw * scale;
            double lx = cx - b / 2, ty = cy - d / 2;

            var pf = new PathFigure { IsClosed = true, StartPoint = new Point(lx, ty) };
            pf.Segments.Add(Seg(lx + b, ty));
            pf.Segments.Add(Seg(lx + b, ty + tf));
            pf.Segments.Add(Seg(lx + tw, ty + tf));
            pf.Segments.Add(Seg(lx + tw, ty + d - tf));
            pf.Segments.Add(Seg(lx + b, ty + d - tf));
            pf.Segments.Add(Seg(lx + b, ty + d));
            pf.Segments.Add(Seg(lx, ty + d));

            return new Path { Data = new PathGeometry(new[] { pf }) };
        }

        private static Path BuildAngle(double cx, double cy, double t3Raw, double t2Raw, double tfRaw, double twRaw, ref double scale)
        {
            LimitScale(ref scale, t3Raw, t2Raw, 80);
            double d = t3Raw * scale, b = t2Raw * scale, tf = tfRaw * scale, tw = twRaw * scale;
            double lx = cx - b / 2, ty = cy - d / 2;

            var pf = new PathFigure { IsClosed = true, StartPoint = new Point(lx, ty) };
            pf.Segments.Add(Seg(lx + tw, ty));
            pf.Segments.Add(Seg(lx + tw, ty + d - tf));
            pf.Segments.Add(Seg(lx + b, ty + d - tf));
            pf.Segments.Add(Seg(lx + b, ty + d));
            pf.Segments.Add(Seg(lx, ty + d));

            return new Path { Data = new PathGeometry(new[] { pf }) };
        }

        // --- Helpers ---

        private static void DrawPreviewGrid(Canvas canvas)
        {
            const double size = 100;
            const double step = 5;
            var minorBrush = new SolidColorBrush(Color.FromRgb(150, 184, 190));
            var majorBrush = new SolidColorBrush(Color.FromRgb(92, 134, 145));
            var xAxisBrush = new SolidColorBrush(Color.FromRgb(210, 55, 55));
            var yAxisBrush = new SolidColorBrush(Color.FromRgb(55, 85, 220));
            var dash = new DoubleCollection { 2, 2 };

            for (double p = 0; p <= size; p += step)
            {
                bool isMajor = Math.Abs(p % 20) < 0.001;
                Brush brush = isMajor ? majorBrush : minorBrush;
                double thickness = isMajor ? 0.55 : 0.35;

                AddGridLine(canvas, p, 0, p, size, brush, thickness, dash);
                AddGridLine(canvas, 0, p, size, p, brush, thickness, dash);
            }

            AddGridLine(canvas, 50, 0, 50, size, yAxisBrush, 0.9, null);
            AddGridLine(canvas, 0, 50, size, 50, xAxisBrush, 0.9, null);
        }

        private static void AddGridLine(Canvas canvas, double x1, double y1, double x2, double y2, Brush brush, double thickness, DoubleCollection dash)
        {
            var line = new Line
            {
                X1 = x1,
                Y1 = y1,
                X2 = x2,
                Y2 = y2,
                Stroke = brush,
                StrokeThickness = thickness,
                SnapsToDevicePixels = true
            };

            if (dash != null)
            {
                line.StrokeDashArray = dash;
            }

            canvas.Children.Add(line);
        }

        private static LineSegment Seg(double x, double y) => new LineSegment(new Point(x, y), true);

        private static (double d, double b) Rescale(double t3, double t2, ref double scale, double limit)
        {
            LimitScale(ref scale, t3, t2, limit);
            return (t3 * scale, t2 * scale);
        }

        private static void LimitScale(ref double scale, double t3, double t2, double limit)
        {
            double maxDim = Math.Max(t3, t2);
            if (maxDim * scale > limit)
                scale = limit / maxDim;
        }

        private static double Get(System.Collections.Generic.Dictionary<string, double> d, string key, double fallback = 0)
        {
            return d.ContainsKey(key) ? d[key] : fallback;
        }

        private static double Get(System.Collections.Generic.Dictionary<string, double> d,
            string key1, string key2, double fallback = 0)
        {
            if (d.ContainsKey(key1)) return d[key1];
            if (d.ContainsKey(key2)) return d[key2];
            return fallback;
        }

        private static double Get(System.Collections.Generic.Dictionary<string, double> d, double fallback, params string[] keys)
        {
            foreach (string key in keys)
            {
                if (d.ContainsKey(key)) return d[key];
            }

            return fallback;
        }

        private static void AddLabel(Canvas canvas, string text)
        {
            var tb = new TextBlock
            {
                Text = text,
                Foreground = new SolidColorBrush(Color.FromRgb(120, 130, 150)),
                FontSize = 8,
                FontFamily = new FontFamily("Segoe UI"),
                TextAlignment = TextAlignment.Center,
                Width = 80
            };
            Canvas.SetLeft(tb, 10);
            Canvas.SetTop(tb, 42);
            canvas.Children.Add(tb);
        }
    }
}
