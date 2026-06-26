using System.Drawing;
using System.Drawing.Drawing2D;

namespace ExcelCSIToolBoxAddIn.AddIn.Ribbon
{
    internal static class PostprocessingRibbonIcons
    {
        internal static readonly Image BaseReactions = CreateBaseReactions();
        internal static readonly Image ModalMassParticipation = CreateModalMassParticipation();
        internal static readonly Image StoryForces = CreateStoryForces();
        internal static readonly Image StoryDrifts = CreateStoryDrifts();
        internal static readonly Image StoryMaxOverAverageDisplacements = CreateMaxOverAverageDisplacements();
        internal static readonly Image StoryMaxOverAverageDrifts = CreateMaxOverAverageDrifts();

        private static Bitmap CreateBaseReactions()
        {
            return CreateIcon(graphics =>
            {
                using (var pen = CreatePen())
                {
                    graphics.DrawRectangle(pen, 26, 7, 10, 24);
                    graphics.DrawLine(pen, 19, 33, 43, 33);
                    graphics.DrawLine(pen, 22, 33, 22, 43);
                    graphics.DrawLine(pen, 29, 33, 29, 43);
                    graphics.DrawLine(pen, 36, 33, 36, 43);
                    DrawArrow(graphics, pen, new Point(22, 45), new Point(22, 52));
                    DrawArrow(graphics, pen, new Point(29, 45), new Point(29, 52));
                    DrawArrow(graphics, pen, new Point(36, 45), new Point(36, 52));
                    graphics.DrawArc(pen, 39, 19, 16, 16, 270, 95);
                    DrawArrow(graphics, pen, new Point(49, 20), new Point(52, 28));
                }
            });
        }

        private static Bitmap CreateModalMassParticipation()
        {
            return CreateIcon(graphics =>
            {
                using (var pen = CreatePen())
                {
                    DrawBuilding(graphics, pen, 13, 8, 27, 49, 5);
                    graphics.DrawLine(pen, 44, 48, 44, 39);
                    graphics.DrawLine(pen, 50, 48, 50, 31);
                    graphics.DrawLine(pen, 56, 48, 56, 23);
                    graphics.DrawLine(pen, 41, 48, 59, 48);
                }
            });
        }

        private static Bitmap CreateStoryForces()
        {
            return CreateIcon(graphics =>
            {
                using (var pen = CreatePen())
                {
                    DrawBuilding(graphics, pen, 14, 8, 27, 49, 5);
                    DrawArrow(graphics, pen, new Point(43, 16), new Point(56, 16));
                    DrawArrow(graphics, pen, new Point(43, 27), new Point(56, 27));
                    DrawArrow(graphics, pen, new Point(43, 38), new Point(56, 38));
                }
            });
        }

        private static Bitmap CreateStoryDrifts()
        {
            return CreateIcon(graphics =>
            {
                using (var pen = CreatePen())
                {
                    Point[] left = { new Point(21, 50), new Point(24, 39), new Point(27, 28), new Point(30, 17), new Point(33, 7) };
                    Point[] right = { new Point(42, 50), new Point(45, 39), new Point(48, 28), new Point(51, 17), new Point(54, 7) };
                    graphics.DrawLines(pen, left);
                    graphics.DrawLines(pen, right);
                    for (int index = 0; index < 5; index++) graphics.DrawLine(pen, left[index], right[index]);
                    graphics.DrawLine(pen, 16, 50, 47, 50);
                    graphics.DrawLine(pen, 24, 6, 56, 6);
                }
            });
        }

        private static Bitmap CreateMaxOverAverageDisplacements()
        {
            return CreateIcon(graphics =>
            {
                using (var pen = CreatePen())
                {
                    DrawBuilding(graphics, pen, 13, 8, 27, 49, 5);
                    DrawArrow(graphics, pen, new Point(43, 16), new Point(57, 16));
                    DrawArrow(graphics, pen, new Point(43, 38), new Point(53, 38));
                    graphics.DrawLine(pen, 45, 27, 58, 27);
                    graphics.DrawLine(pen, 45, 24, 45, 30);
                    graphics.DrawLine(pen, 58, 24, 58, 30);
                }
            });
        }

        private static Bitmap CreateMaxOverAverageDrifts()
        {
            return CreateIcon(graphics =>
            {
                using (var pen = CreatePen())
                {
                    Point[] left = { new Point(18, 50), new Point(22, 39), new Point(20, 28), new Point(27, 17), new Point(25, 7) };
                    Point[] right = { new Point(39, 50), new Point(43, 39), new Point(41, 28), new Point(48, 17), new Point(46, 7) };
                    graphics.DrawLines(pen, left);
                    graphics.DrawLines(pen, right);
                    for (int index = 0; index < 5; index++) graphics.DrawLine(pen, left[index], right[index]);
                    graphics.DrawLine(pen, 13, 50, 44, 50);
                    DrawArrow(graphics, pen, new Point(51, 15), new Point(60, 15));
                    DrawArrow(graphics, pen, new Point(51, 36), new Point(58, 36));
                }
            });
        }

        private static Bitmap CreateIcon(System.Action<Graphics> draw)
        {
            var bitmap = new Bitmap(64, 64);
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.SmoothingMode = SmoothingMode.AntiAlias;
                graphics.PixelOffsetMode = PixelOffsetMode.HighQuality;
                draw(graphics);
            }

            return bitmap;
        }

        private static Pen CreatePen()
        {
            return new Pen(Color.White, 2.6f)
            {
                StartCap = LineCap.Round,
                EndCap = LineCap.Round,
                LineJoin = LineJoin.Round
            };
        }

        private static void DrawBuilding(Graphics graphics, Pen pen, int left, int top, int width, int height, int floors)
        {
            graphics.DrawRectangle(pen, left, top, width, height);
            int floorHeight = height / floors;
            for (int floor = 1; floor < floors; floor++) graphics.DrawLine(pen, left, top + floor * floorHeight, left + width, top + floor * floorHeight);
            graphics.DrawLine(pen, left - 4, top + height, left + width + 4, top + height);
        }

        private static void DrawArrow(Graphics graphics, Pen pen, Point start, Point end)
        {
            graphics.DrawLine(pen, start, end);
            float angle = (float)System.Math.Atan2(end.Y - start.Y, end.X - start.X);
            const float arrowLength = 5f;
            const float arrowAngle = 0.6f;
            PointF first = new PointF(end.X - arrowLength * (float)System.Math.Cos(angle - arrowAngle), end.Y - arrowLength * (float)System.Math.Sin(angle - arrowAngle));
            PointF second = new PointF(end.X - arrowLength * (float)System.Math.Cos(angle + arrowAngle), end.Y - arrowLength * (float)System.Math.Sin(angle + arrowAngle));
            graphics.DrawLine(pen, end, first);
            graphics.DrawLine(pen, end, second);
        }
    }
}
