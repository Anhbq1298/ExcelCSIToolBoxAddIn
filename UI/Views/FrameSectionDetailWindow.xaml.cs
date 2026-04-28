using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Shapes;
using ExcelCSIToolBoxAddIn.Data.DTOs;

namespace ExcelCSIToolBoxAddIn.UI.Views
{
    public partial class FrameSectionDetailWindow : Window
    {
        public FrameSectionDetailWindow(object dataContext)
        {
            InitializeComponent();
            DataContext = dataContext;
            Loaded += FrameSectionDetailWindow_Loaded;
        }

        private void FrameSectionDetailWindow_Loaded(object sender, RoutedEventArgs e)
        {
            if (DataContext is CSISapModelFrameSectionDetailDTO dto)
            {
                DrawSectionShape(dto);
            }
        }

        private void DrawSectionShape(CSISapModelFrameSectionDetailDTO dto)
        {
            PreviewCanvas.Children.Clear();
            var strokeBrush = new SolidColorBrush(Colors.Black);
            var fillBrush = new SolidColorBrush(Color.FromArgb(100, 100, 149, 237));
            double strokeThickness = 1.0;

            // Define the bounding box to be 100x100
            double cx = 50;
            double cy = 50;

            double t3 = dto.Dimensions.ContainsKey("Total depth ( t3 )") ? dto.Dimensions["Total depth ( t3 )"] : 
                        dto.Dimensions.ContainsKey("Depth ( t3 )") ? dto.Dimensions["Depth ( t3 )"] : 
                        dto.Dimensions.ContainsKey("Outside diameter ( t3 )") ? dto.Dimensions["Outside diameter ( t3 )"] : 
                        dto.Dimensions.ContainsKey("Diameter ( t3 )") ? dto.Dimensions["Diameter ( t3 )"] : 100;

            double scale = t3 > 0 ? 80.0 / t3 : 1.0;

            Path shapePath = new Path
            {
                Stroke = strokeBrush,
                StrokeThickness = strokeThickness,
                Fill = fillBrush
            };

            GeometryGroup group = new GeometryGroup();

            switch (dto.ShapeType)
            {
                case FrameSectionShapeType.Pipe:
                {
                    double d = t3 * scale;
                    double tw = (dto.Dimensions.ContainsKey("Wall thickness ( tw )") ? dto.Dimensions["Wall thickness ( tw )"] : 0) * scale;
                    
                    group.Children.Add(new EllipseGeometry(new Point(cx, cy), d / 2, d / 2));
                    if (tw > 0 && d / 2 - tw > 0)
                    {
                        group.Children.Add(new EllipseGeometry(new Point(cx, cy), d / 2 - tw, d / 2 - tw));
                    }
                    shapePath.Data = group;
                    break;
                }
                case FrameSectionShapeType.Circular:
                {
                    double d = t3 * scale;
                    group.Children.Add(new EllipseGeometry(new Point(cx, cy), d / 2, d / 2));
                    shapePath.Data = group;
                    break;
                }
                case FrameSectionShapeType.Rectangular:
                {
                    double d = t3 * scale;
                    double b = (dto.Dimensions.ContainsKey("Width ( t2 )") ? dto.Dimensions["Width ( t2 )"] : t3) * scale;
                    // Adjust scale if width is much larger than depth
                    if (b > 80)
                    {
                        scale = scale * (80 / b);
                        d = t3 * scale;
                        b = (dto.Dimensions.ContainsKey("Width ( t2 )") ? dto.Dimensions["Width ( t2 )"] : t3) * scale;
                    }

                    group.Children.Add(new RectangleGeometry(new Rect(cx - b / 2, cy - d / 2, b, d)));
                    shapePath.Data = group;
                    break;
                }
                case FrameSectionShapeType.Tube:
                {
                    double d = t3 * scale;
                    double b = (dto.Dimensions.ContainsKey("Flange width ( t2 )") ? dto.Dimensions["Flange width ( t2 )"] : t3) * scale;
                    double tf = (dto.Dimensions.ContainsKey("Flange thickness ( tf )") ? dto.Dimensions["Flange thickness ( tf )"] : 0) * scale;
                    double tw = (dto.Dimensions.ContainsKey("Web thickness ( tw )") ? dto.Dimensions["Web thickness ( tw )"] : 0) * scale;

                    if (b > 80)
                    {
                        scale = scale * (80 / b);
                        d = t3 * scale;
                        b = (dto.Dimensions.ContainsKey("Flange width ( t2 )") ? dto.Dimensions["Flange width ( t2 )"] : t3) * scale;
                        tf = (dto.Dimensions.ContainsKey("Flange thickness ( tf )") ? dto.Dimensions["Flange thickness ( tf )"] : 0) * scale;
                        tw = (dto.Dimensions.ContainsKey("Web thickness ( tw )") ? dto.Dimensions["Web thickness ( tw )"] : 0) * scale;
                    }

                    group.Children.Add(new RectangleGeometry(new Rect(cx - b / 2, cy - d / 2, b, d)));
                    if (b - 2 * tw > 0 && d - 2 * tf > 0)
                    {
                        group.Children.Add(new RectangleGeometry(new Rect(cx - b / 2 + tw, cy - d / 2 + tf, b - 2 * tw, d - 2 * tf)));
                    }
                    shapePath.Data = group;
                    break;
                }
                case FrameSectionShapeType.I:
                {
                    double d = t3 * scale;
                    double t2 = (dto.Dimensions.ContainsKey("Top flange width ( t2 )") ? dto.Dimensions["Top flange width ( t2 )"] : t3) * scale;
                    double tf = (dto.Dimensions.ContainsKey("Top flange thickness ( tf )") ? dto.Dimensions["Top flange thickness ( tf )"] : 0) * scale;
                    double tw = (dto.Dimensions.ContainsKey("Web thickness ( tw )") ? dto.Dimensions["Web thickness ( tw )"] : 0) * scale;
                    double t2b = (dto.Dimensions.ContainsKey("Bottom flange width ( t2b )") ? dto.Dimensions["Bottom flange width ( t2b )"] : t2) * scale;
                    double tfb = (dto.Dimensions.ContainsKey("Bottom flange thickness ( tfb )") ? dto.Dimensions["Bottom flange thickness ( tfb )"] : tf) * scale;

                    double maxB = Math.Max(t2, t2b);
                    if (maxB > 80)
                    {
                        scale = scale * (80 / maxB);
                        d = t3 * scale;
                        t2 *= scale; tf *= scale; tw *= scale; t2b *= scale; tfb *= scale;
                    }

                    PathFigure pf = new PathFigure { StartPoint = new Point(cx - t2 / 2, cy - d / 2), IsClosed = true };
                    pf.Segments.Add(new LineSegment(new Point(cx + t2 / 2, cy - d / 2), true));
                    pf.Segments.Add(new LineSegment(new Point(cx + t2 / 2, cy - d / 2 + tf), true));
                    pf.Segments.Add(new LineSegment(new Point(cx + tw / 2, cy - d / 2 + tf), true));
                    pf.Segments.Add(new LineSegment(new Point(cx + tw / 2, cy + d / 2 - tfb), true));
                    pf.Segments.Add(new LineSegment(new Point(cx + t2b / 2, cy + d / 2 - tfb), true));
                    pf.Segments.Add(new LineSegment(new Point(cx + t2b / 2, cy + d / 2), true));
                    pf.Segments.Add(new LineSegment(new Point(cx - t2b / 2, cy + d / 2), true));
                    pf.Segments.Add(new LineSegment(new Point(cx - t2b / 2, cy + d / 2 - tfb), true));
                    pf.Segments.Add(new LineSegment(new Point(cx - tw / 2, cy + d / 2 - tfb), true));
                    pf.Segments.Add(new LineSegment(new Point(cx - tw / 2, cy - d / 2 + tf), true));
                    pf.Segments.Add(new LineSegment(new Point(cx - t2 / 2, cy - d / 2 + tf), true));
                    
                    PathGeometry pg = new PathGeometry();
                    pg.Figures.Add(pf);
                    group.Children.Add(pg);
                    shapePath.Data = group;
                    break;
                }
                case FrameSectionShapeType.Channel:
                {
                    double d = t3 * scale;
                    double b = (dto.Dimensions.ContainsKey("Flange width ( t2 )") ? dto.Dimensions["Flange width ( t2 )"] : t3) * scale;
                    double tf = (dto.Dimensions.ContainsKey("Flange thickness ( tf )") ? dto.Dimensions["Flange thickness ( tf )"] : 0) * scale;
                    double tw = (dto.Dimensions.ContainsKey("Web thickness ( tw )") ? dto.Dimensions["Web thickness ( tw )"] : 0) * scale;

                    if (b > 80)
                    {
                        scale = scale * (80 / b);
                        d *= scale; b *= scale; tf *= scale; tw *= scale;
                    }

                    // Assuming Channel opening to the right, cg somewhat in the middle. Let's align left edge.
                    double leftX = cx - b / 2;
                    double topY = cy - d / 2;

                    PathFigure pf = new PathFigure { StartPoint = new Point(leftX, topY), IsClosed = true };
                    pf.Segments.Add(new LineSegment(new Point(leftX + b, topY), true));
                    pf.Segments.Add(new LineSegment(new Point(leftX + b, topY + tf), true));
                    pf.Segments.Add(new LineSegment(new Point(leftX + tw, topY + tf), true));
                    pf.Segments.Add(new LineSegment(new Point(leftX + tw, topY + d - tf), true));
                    pf.Segments.Add(new LineSegment(new Point(leftX + b, topY + d - tf), true));
                    pf.Segments.Add(new LineSegment(new Point(leftX + b, topY + d), true));
                    pf.Segments.Add(new LineSegment(new Point(leftX, topY + d), true));

                    PathGeometry pg = new PathGeometry();
                    pg.Figures.Add(pf);
                    group.Children.Add(pg);
                    shapePath.Data = group;
                    break;
                }
                case FrameSectionShapeType.Angle:
                {
                    double d = t3 * scale;
                    double b = (dto.Dimensions.ContainsKey("Flange width ( t2 )") ? dto.Dimensions["Flange width ( t2 )"] : t3) * scale;
                    double tf = (dto.Dimensions.ContainsKey("Flange thickness ( tf )") ? dto.Dimensions["Flange thickness ( tf )"] : 0) * scale;
                    double tw = (dto.Dimensions.ContainsKey("Web thickness ( tw )") ? dto.Dimensions["Web thickness ( tw )"] : 0) * scale;

                    if (b > 80)
                    {
                        scale = scale * (80 / b);
                        d *= scale; b *= scale; tf *= scale; tw *= scale;
                    }

                    double leftX = cx - b / 2;
                    double topY = cy - d / 2;

                    PathFigure pf = new PathFigure { StartPoint = new Point(leftX, topY), IsClosed = true };
                    pf.Segments.Add(new LineSegment(new Point(leftX + tw, topY), true));
                    pf.Segments.Add(new LineSegment(new Point(leftX + tw, topY + d - tf), true));
                    pf.Segments.Add(new LineSegment(new Point(leftX + b, topY + d - tf), true));
                    pf.Segments.Add(new LineSegment(new Point(leftX + b, topY + d), true));
                    pf.Segments.Add(new LineSegment(new Point(leftX, topY + d), true));

                    PathGeometry pg = new PathGeometry();
                    pg.Figures.Add(pf);
                    group.Children.Add(pg);
                    shapePath.Data = group;
                    break;
                }
                default:
                {
                    TextBlock tb = new TextBlock
                    {
                        Text = "No Preview",
                        Foreground = new SolidColorBrush(Colors.Gray),
                        HorizontalAlignment = HorizontalAlignment.Center,
                        VerticalAlignment = VerticalAlignment.Center
                    };
                    Canvas.SetLeft(tb, 20);
                    Canvas.SetTop(tb, 40);
                    PreviewCanvas.Children.Add(tb);
                    return;
                }
            }

            PreviewCanvas.Children.Add(shapePath);
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
