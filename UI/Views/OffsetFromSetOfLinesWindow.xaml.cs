using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Shapes;
using ExcelCSIToolBox.Application.Modelling.OffsetPolylines;
using ExcelCSIToolBoxAddIn.UI.ViewModels;

namespace ExcelCSIToolBoxAddIn.UI.Views
{
    public partial class OffsetFromSetOfLinesWindow : Window
    {
        private readonly CsiToolboxViewModel _viewModel;

        public OffsetFromSetOfLinesWindow(CsiToolboxViewModel viewModel)
        {
            InitializeComponent();

            _viewModel = viewModel ?? throw new ArgumentNullException(nameof(viewModel));
            DataContext = _viewModel;
            _viewModel.PropertyChanged += OnViewModelPropertyChanged;
            OffsetPolylinePreviewCanvas.SizeChanged += OffsetPolylinePreviewCanvas_SizeChanged;
            Closed += OnClosed;
            RenderOffsetPolylinePreview(_viewModel.OffsetPreviewResult);
        }

        private void OnClosed(object sender, EventArgs e)
        {
            _viewModel.PropertyChanged -= OnViewModelPropertyChanged;
            OffsetPolylinePreviewCanvas.SizeChanged -= OffsetPolylinePreviewCanvas_SizeChanged;
        }

        private void OnViewModelPropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == nameof(CsiToolboxViewModel.OffsetPreviewVersion))
            {
                RenderOffsetPolylinePreview(_viewModel.OffsetPreviewResult);
            }
        }

        private void OffsetPolylinePreviewCanvas_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            RenderOffsetPolylinePreview(_viewModel.OffsetPreviewResult);
        }

        private void RenderOffsetPolylinePreview(OffsetPolylineResult result)
        {
            if (OffsetPolylinePreviewCanvas == null)
            {
                return;
            }

            OffsetPolylinePreviewCanvas.Children.Clear();
            if (result == null || result.OriginalVertices == null || result.OriginalVertices.Count < 3)
            {
                return;
            }

            List<Point> original = ProjectPreviewPoints(result, result.OriginalVertices);
            List<Point> offset = result.ResultVertices == null
                ? new List<Point>()
                : ProjectPreviewPoints(result, result.ResultVertices);

            Rect bounds = GetBounds(original, offset);
            if (bounds.Width <= 0 || bounds.Height <= 0)
            {
                return;
            }

            double canvasWidth = OffsetPolylinePreviewCanvas.ActualWidth > 20 ? OffsetPolylinePreviewCanvas.ActualWidth : 520;
            double canvasHeight = OffsetPolylinePreviewCanvas.ActualHeight > 20 ? OffsetPolylinePreviewCanvas.ActualHeight : 360;
            const double padding = 24;
            double scale = Math.Min(
                (canvasWidth - padding * 2) / bounds.Width,
                (canvasHeight - padding * 2) / bounds.Height);
            if (double.IsInfinity(scale) || double.IsNaN(scale) || scale <= 0)
            {
                scale = 1;
            }

            Func<Point, Point> map = delegate (Point point)
            {
                double x = padding + (point.X - bounds.Left) * scale;
                double y = canvasHeight - padding - (point.Y - bounds.Top) * scale;
                return new Point(x, y);
            };

            DrawClosedPolyline(original, map, new SolidColorBrush(Color.FromRgb(128, 138, 148)), 1.5);
            DrawVertices(original, map, new SolidColorBrush(Color.FromRgb(128, 138, 148)), 3.5);
            DrawSegmentLabels(original, map, new SolidColorBrush(Color.FromRgb(100, 116, 139)), "S");

            if (offset.Count >= 3)
            {
                Brush resultBrush = result.OffsetDistance < 0
                    ? new SolidColorBrush(Color.FromRgb(190, 92, 35))
                    : new SolidColorBrush(Color.FromRgb(47, 128, 237));
                DrawClosedPolyline(offset, map, resultBrush, 2.2);
                DrawVertices(offset, map, resultBrush, 4.2);
                DrawSegmentLabels(offset, map, resultBrush, "R");
            }
        }

        private void DrawClosedPolyline(IReadOnlyList<Point> points, Func<Point, Point> map, Brush stroke, double thickness)
        {
            if (points == null || points.Count < 2)
            {
                return;
            }

            var pointCollection = new PointCollection();
            for (int i = 0; i < points.Count; i++)
            {
                pointCollection.Add(map(points[i]));
            }

            pointCollection.Add(map(points[0]));
            OffsetPolylinePreviewCanvas.Children.Add(new Polyline
            {
                Points = pointCollection,
                Stroke = stroke,
                StrokeThickness = thickness,
                StrokeLineJoin = PenLineJoin.Miter
            });
        }

        private void DrawVertices(IReadOnlyList<Point> points, Func<Point, Point> map, Brush brush, double radius)
        {
            for (int i = 0; i < points.Count; i++)
            {
                Point point = map(points[i]);
                var ellipse = new Ellipse
                {
                    Width = radius * 2,
                    Height = radius * 2,
                    Fill = brush,
                    Stroke = Brushes.White,
                    StrokeThickness = 1
                };
                Canvas.SetLeft(ellipse, point.X - radius);
                Canvas.SetTop(ellipse, point.Y - radius);
                OffsetPolylinePreviewCanvas.Children.Add(ellipse);
            }
        }

        private void DrawSegmentLabels(IReadOnlyList<Point> points, Func<Point, Point> map, Brush brush, string prefix)
        {
            for (int i = 0; i < points.Count; i++)
            {
                Point start = map(points[i]);
                Point end = map(points[(i + 1) % points.Count]);
                var label = new System.Windows.Controls.TextBlock
                {
                    Text = prefix + (i + 1).ToString(System.Globalization.CultureInfo.CurrentCulture),
                    FontSize = 9,
                    Foreground = brush,
                    Background = Brushes.White,
                    Padding = new Thickness(2, 0, 2, 0)
                };
                Canvas.SetLeft(label, (start.X + end.X) * 0.5 + 3);
                Canvas.SetTop(label, (start.Y + end.Y) * 0.5 + 3);
                OffsetPolylinePreviewCanvas.Children.Add(label);
            }
        }

        private static List<Point> ProjectPreviewPoints(OffsetPolylineResult result, IReadOnlyList<OffsetPoint3D> points)
        {
            var projected = new List<Point>();
            foreach (OffsetPoint3D point in points)
            {
                OffsetPoint3D relative = new OffsetPoint3D(
                    point.X - result.PlaneOrigin.X,
                    point.Y - result.PlaneOrigin.Y,
                    point.Z - result.PlaneOrigin.Z);
                double u = Dot(relative, result.PlaneXAxis);
                double v = Dot(relative, result.PlaneYAxis);
                projected.Add(new Point(u, v));
            }

            return projected;
        }

        private static Rect GetBounds(IReadOnlyList<Point> original, IReadOnlyList<Point> offset)
        {
            bool initialized = false;
            double minX = 0;
            double maxX = 0;
            double minY = 0;
            double maxY = 0;
            Action<Point> addPoint = delegate (Point point)
            {
                if (!initialized)
                {
                    minX = maxX = point.X;
                    minY = maxY = point.Y;
                    initialized = true;
                    return;
                }

                minX = Math.Min(minX, point.X);
                maxX = Math.Max(maxX, point.X);
                minY = Math.Min(minY, point.Y);
                maxY = Math.Max(maxY, point.Y);
            };

            foreach (Point point in original) addPoint(point);
            foreach (Point point in offset) addPoint(point);

            return initialized
                ? new Rect(minX, minY, Math.Max(1, maxX - minX), Math.Max(1, maxY - minY))
                : Rect.Empty;
        }

        private static double Dot(OffsetPoint3D left, OffsetPoint3D right)
        {
            return left.X * right.X + left.Y * right.Y + left.Z * right.Z;
        }
    }
}
