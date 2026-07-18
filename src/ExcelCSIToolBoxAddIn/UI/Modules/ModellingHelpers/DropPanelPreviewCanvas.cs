using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Windows;
using System.Windows.Media;
using ExcelCSIToolBox.Application.Modelling.DropPanels;

namespace ExcelCSIToolBoxAddIn.UI.Views
{
    public sealed class DropPanelPreviewCanvas : FrameworkElement
    {
        public static readonly DependencyProperty PlanProperty = DependencyProperty.Register(
            "Plan",
            typeof(DropPanelOperationPlan),
            typeof(DropPanelPreviewCanvas),
            new FrameworkPropertyMetadata(null, FrameworkPropertyMetadataOptions.AffectsRender));

        public DropPanelOperationPlan Plan
        {
            get { return (DropPanelOperationPlan)GetValue(PlanProperty); }
            set { SetValue(PlanProperty, value); }
        }

        protected override void OnRender(DrawingContext drawingContext)
        {
            base.OnRender(drawingContext);
            drawingContext.DrawRectangle(Brushes.White, null, new Rect(0.0, 0.0, ActualWidth, ActualHeight));
            if (Plan == null)
            {
                DrawCenteredText(drawingContext, "Run Preview to inspect the proposed ETABS regions.", Brushes.Gray);
                return;
            }

            List<DropPanelPoint3D> allPoints = CollectPoints(Plan);
            if (allPoints.Count == 0)
            {
                DrawCenteredText(drawingContext, "The preview contains no drawable geometry.", Brushes.Gray);
                return;
            }

            const double margin = 34.0;
            double minX = allPoints.Min(point => point.X);
            double maxX = allPoints.Max(point => point.X);
            double minY = allPoints.Min(point => point.Y);
            double maxY = allPoints.Max(point => point.Y);
            double width = Math.Max(maxX - minX, 1e-9);
            double height = Math.Max(maxY - minY, 1e-9);
            double scale = Math.Min(
                Math.Max(1.0, ActualWidth - margin * 2.0) / width,
                Math.Max(1.0, ActualHeight - margin * 2.0) / height);

            Func<DropPanelPoint3D, Point> map = point => new Point(
                margin + (point.X - minX) * scale,
                ActualHeight - margin - (point.Y - minY) * scale);

            bool isInvalidPlan = !Plan.IsValid;
            Pen sourcePen = new Pen(
                new SolidColorBrush(isInvalidPlan ? Color.FromRgb(185, 28, 28) : Color.FromRgb(71, 85, 105)),
                isInvalidPlan ? 2.0 : 1.2)
            {
                DashStyle = DashStyles.Dash
            };
            Pen normalPen = new Pen(new SolidColorBrush(Color.FromRgb(37, 99, 235)), 1.0);
            Pen dropPen = new Pen(new SolidColorBrush(Color.FromRgb(180, 83, 9)), 1.3);
            Pen openingPen = new Pen(new SolidColorBrush(Color.FromRgb(185, 28, 28)), 1.4);
            Pen requestPen = new Pen(new SolidColorBrush(Color.FromRgb(234, 88, 12)), 1.0)
            {
                DashStyle = DashStyles.Dot
            };
            Brush normalFill = new SolidColorBrush(Color.FromArgb(72, 147, 197, 253));
            Brush dropFill = new SolidColorBrush(Color.FromArgb(100, 251, 191, 36));
            Brush openingFill = new SolidColorBrush(Color.FromArgb(90, 254, 202, 202));

            foreach (DropPanelAreaInfo source in Plan.SourceAreas)
            {
                DrawPolygon(drawingContext, source.Points, map, null, sourcePen);
                DrawLabel(drawingContext, source.AreaName, Centroid(source.Points, map), Brushes.DarkSlateGray);
            }

            foreach (DropPanelRegion region in Plan.Regions.Where(item => !item.IsDrop))
            {
                DrawPolygon(drawingContext, region.Points, map, normalFill, normalPen);
                DrawLabel(drawingContext, "Normal", Centroid(region.Points, map), Brushes.Navy);
            }

            foreach (DropPanelRegion region in Plan.Regions.Where(item => item.IsDrop))
            {
                DrawPolygon(drawingContext, region.Points, map, dropFill, dropPen);
                DrawLabel(drawingContext, "Drop", Centroid(region.Points, map), Brushes.DarkGoldenrod);
            }

            foreach (DropPanelAreaInfo opening in Plan.Openings)
            {
                DrawPolygon(drawingContext, opening.Points, map, openingFill, openingPen);
                DrawLabel(drawingContext, "Opening", Centroid(opening.Points, map), Brushes.DarkRed);
            }

            foreach (DropPanelRequest request in Plan.Requests)
            {
                DrawPolygon(drawingContext, request.Points, map, null, requestPen);
            }

            foreach (DropPanelColumnInfo column in Plan.Columns.Where(item => item.IsValid))
            {
                Point point = map(new DropPanelPoint3D(column.X, column.Y, column.Z));
                drawingContext.DrawEllipse(Brushes.DarkRed, new Pen(Brushes.White, 1.0), point, 4.0, 4.0);
                DrawLabel(drawingContext, column.FrameName, new Point(point.X + 6.0, point.Y - 12.0), Brushes.DarkRed);
            }

            DrawLegend(drawingContext, isInvalidPlan);
        }

        private static List<DropPanelPoint3D> CollectPoints(DropPanelOperationPlan plan)
        {
            List<DropPanelPoint3D> points = new List<DropPanelPoint3D>();
            points.AddRange(plan.SourceAreas.SelectMany(area => area.Points));
            points.AddRange(plan.Openings.SelectMany(area => area.Points));
            points.AddRange(plan.Requests.SelectMany(request => request.Points));
            points.AddRange(plan.Regions.SelectMany(region => region.Points));
            return points;
        }

        private static void DrawPolygon(
            DrawingContext drawingContext,
            IReadOnlyList<DropPanelPoint3D> points,
            Func<DropPanelPoint3D, Point> map,
            Brush fill,
            Pen pen)
        {
            if (points == null || points.Count < 3)
            {
                return;
            }

            StreamGeometry geometry = new StreamGeometry();
            using (StreamGeometryContext context = geometry.Open())
            {
                context.BeginFigure(map(points[0]), fill != null, true);
                context.PolyLineTo(points.Skip(1).Select(map).ToList(), true, false);
            }

            geometry.Freeze();
            drawingContext.DrawGeometry(fill, pen, geometry);
        }

        private static Point Centroid(IReadOnlyList<DropPanelPoint3D> points, Func<DropPanelPoint3D, Point> map)
        {
            if (points == null || points.Count == 0)
            {
                return new Point();
            }

            return new Point(points.Average(point => map(point).X), points.Average(point => map(point).Y));
        }

        private static void DrawLabel(DrawingContext drawingContext, string text, Point point, Brush brush)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return;
            }

            FormattedText formatted = CreateText(text, 10.0, brush);
            drawingContext.DrawText(formatted, new Point(point.X - formatted.Width / 2.0, point.Y - formatted.Height / 2.0));
        }

        private void DrawCenteredText(DrawingContext drawingContext, string text, Brush brush)
        {
            FormattedText formatted = CreateText(text, 13.0, brush);
            drawingContext.DrawText(
                formatted,
                new Point(Math.Max(0.0, (ActualWidth - formatted.Width) / 2.0), Math.Max(0.0, (ActualHeight - formatted.Height) / 2.0)));
        }

        private static void DrawLegend(DrawingContext drawingContext, bool isInvalidPlan)
        {
            string legend = "Blue: normal slab   Gold: drop   Red: opening / column   Dotted: requested drop";
            drawingContext.DrawText(CreateText(legend, 10.0, Brushes.DimGray), new Point(8.0, 8.0));
            if (isInvalidPlan)
            {
                drawingContext.DrawText(
                    CreateText("INVALID PREVIEW - Apply is disabled; red dashed source boundaries require review.", 11.0, Brushes.DarkRed),
                    new Point(8.0, 24.0));
            }
        }

        private static FormattedText CreateText(string text, double size, Brush brush)
        {
            return new FormattedText(
                text,
                CultureInfo.CurrentCulture,
                FlowDirection.LeftToRight,
                new Typeface("Segoe UI"),
                size,
                brush,
                1.0);
        }
    }
}
