using System;
using System.Collections.Generic;

namespace ExcelCSIToolBox.Application.Modelling.PileCaps
{
    public class ThreePileCapBoundaryCalculator
    {
        private const double Tolerance = 0.000001;

        public ThreePileCapGeometry Calculate(
            double pileDiameterMillimeters,
            double edgeDistanceMillimeters,
            double spacingMillimeters)
        {
            if (pileDiameterMillimeters <= 0)
            {
                throw new ArgumentOutOfRangeException("pileDiameterMillimeters", "Pile diameter must be greater than zero.");
            }

            if (edgeDistanceMillimeters < 0)
            {
                throw new ArgumentOutOfRangeException("edgeDistanceMillimeters", "Edge distance cannot be negative.");
            }

            if (spacingMillimeters <= 0)
            {
                throw new ArgumentOutOfRangeException("spacingMillimeters", "Pile spacing must be greater than zero.");
            }

            double monoCapSize = pileDiameterMillimeters + edgeDistanceMillimeters * 2.0;
            double triangleHeight = spacingMillimeters * Math.Sqrt(3.0) / 2.0;

            PileCapPoint2D bottomLeftPile = new PileCapPoint2D(-spacingMillimeters / 2.0, -triangleHeight / 3.0);
            PileCapPoint2D bottomRightPile = new PileCapPoint2D(spacingMillimeters / 2.0, -triangleHeight / 3.0);
            PileCapPoint2D topPile = new PileCapPoint2D(0, 2.0 * triangleHeight / 3.0);
            var pileCenters = new List<PileCapPoint2D>
            {
                bottomLeftPile,
                bottomRightPile,
                topPile
            };

            var temporaryMonoCaps = new List<Rectangle2D>
            {
                CreateMonoCap(bottomLeftPile, monoCapSize),
                CreateMonoCap(bottomRightPile, monoCapSize),
                CreateMonoCap(topPile, monoCapSize)
            };

            Rectangle2D lowerLeftCap = temporaryMonoCaps[0];
            Rectangle2D lowerRightCap = temporaryMonoCaps[1];
            Rectangle2D upperCap = temporaryMonoCaps[2];

            var boundary = new List<PileCapPoint2D>
            {
                new PileCapPoint2D(lowerLeftCap.Left, lowerLeftCap.Bottom),
                new PileCapPoint2D(lowerRightCap.Right, lowerLeftCap.Bottom),
                new PileCapPoint2D(lowerRightCap.Right, lowerLeftCap.Top),
                new PileCapPoint2D(upperCap.Right, upperCap.Top),
                new PileCapPoint2D(upperCap.Left, upperCap.Top),
                new PileCapPoint2D(lowerLeftCap.Left, lowerLeftCap.Top)
            };

            var segments = CreateBoundarySegments(boundary);
            ValidateTemporaryCapsInsideBoundary(temporaryMonoCaps, boundary);

            return new ThreePileCapGeometry(
                pileCenters,
                new PileCapPoint2D(0, 0),
                temporaryMonoCaps,
                boundary,
                segments);
        }

        private static Rectangle2D CreateMonoCap(PileCapPoint2D pileCenter, double monoCapSize)
        {
            double half = monoCapSize / 2.0;
            return new Rectangle2D(
                pileCenter.X - half,
                pileCenter.X + half,
                pileCenter.Y - half,
                pileCenter.Y + half);
        }

        private static List<LineSegment2D> CreateBoundarySegments(IReadOnlyList<PileCapPoint2D> boundary)
        {
            var segments = new List<LineSegment2D>();
            for (int i = 0; i < boundary.Count; i++)
            {
                PileCapPoint2D start = boundary[i];
                PileCapPoint2D end = boundary[(i + 1) % boundary.Count];
                if (Distance(start, end) > Tolerance)
                {
                    segments.Add(new LineSegment2D(start, end));
                }
            }

            if (segments.Count != boundary.Count)
            {
                throw new InvalidOperationException("Three-pile cap boundary contains a duplicate or zero-length segment.");
            }

            return segments;
        }

        private static void ValidateTemporaryCapsInsideBoundary(
            IReadOnlyList<Rectangle2D> temporaryMonoCaps,
            IReadOnlyList<PileCapPoint2D> boundary)
        {
            foreach (Rectangle2D temporaryMonoCap in temporaryMonoCaps)
            {
                foreach (PileCapPoint2D corner in temporaryMonoCap.GetCornersClockwise())
                {
                    if (!IsPointInsideOrOnBoundary(corner, boundary))
                    {
                        throw new InvalidOperationException("Three-pile cap boundary does not contain all temporary mono pile-cap rectangles.");
                    }
                }
            }
        }

        public static bool IsPointInsideOrOnBoundary(
            PileCapPoint2D point,
            IReadOnlyList<PileCapPoint2D> boundary)
        {
            if (boundary == null || boundary.Count < 3)
            {
                return false;
            }

            bool inside = false;
            int j = boundary.Count - 1;
            for (int i = 0; i < boundary.Count; i++)
            {
                PileCapPoint2D a = boundary[j];
                PileCapPoint2D b = boundary[i];

                if (IsPointOnSegment(point, a, b))
                {
                    return true;
                }

                bool crossesRay = ((b.Y > point.Y) != (a.Y > point.Y)) &&
                                  (point.X < (a.X - b.X) * (point.Y - b.Y) / (a.Y - b.Y) + b.X);
                if (crossesRay)
                {
                    inside = !inside;
                }

                j = i;
            }

            return inside;
        }

        private static bool IsPointOnSegment(PileCapPoint2D point, PileCapPoint2D start, PileCapPoint2D end)
        {
            double cross = (point.Y - start.Y) * (end.X - start.X) -
                           (point.X - start.X) * (end.Y - start.Y);
            if (Math.Abs(cross) > Tolerance)
            {
                return false;
            }

            double dot = (point.X - start.X) * (end.X - start.X) +
                         (point.Y - start.Y) * (end.Y - start.Y);
            if (dot < -Tolerance)
            {
                return false;
            }

            double lengthSquared = (end.X - start.X) * (end.X - start.X) +
                                   (end.Y - start.Y) * (end.Y - start.Y);
            return dot <= lengthSquared + Tolerance;
        }

        private static double Distance(PileCapPoint2D left, PileCapPoint2D right)
        {
            double dx = left.X - right.X;
            double dy = left.Y - right.Y;
            return Math.Sqrt(dx * dx + dy * dy);
        }
    }
}
