using System;
using System.Collections.Generic;
using ExcelCSIToolBox.Core.Contracts.CSI.PileCap;

namespace ExcelCSIToolBox.Application.Modelling.PileCaps
{
    public class PileCapGeometryCalculator
    {
        private const double Tolerance = 0.000001;

        public PileCapGeometry Calculate(PileCapInputParameters input)
        {
            if (input == null)
            {
                throw new ArgumentNullException("input");
            }

            PileCapInputParameters normalized = NormalizeSpacing(input);
            switch (normalized.ArrangementType)
            {
                case PileCapArrangementType.Mono:
                    return CalculateMonoPileLayout(normalized);
                case PileCapArrangementType.TwoPile:
                    return CalculateTwoPileLayout(normalized);
                case PileCapArrangementType.ThreePile:
                    return CalculateThreePileLayout(normalized);
                case PileCapArrangementType.FourPile:
                    return CalculateFourPileLayout(normalized);
                default:
                    throw new ArgumentOutOfRangeException("input", "Unsupported pile-cap arrangement.");
            }
        }

        public IReadOnlyList<PileCapPoint2D> RotatePoints(
            IReadOnlyList<PileCapPoint2D> points,
            double rotationDegrees)
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

        private static PileCapInputParameters NormalizeSpacing(PileCapInputParameters input)
        {
            var normalized = new PileCapInputParameters
            {
                ArrangementType = input.ArrangementType,
                PileDiameterMillimeters = input.PileDiameterMillimeters,
                PileLengthMillimeters = input.PileLengthMillimeters,
                RotationDegrees = input.RotationDegrees,
                AutoSpacing = input.AutoSpacing,
                PileSpacingMillimeters = input.PileSpacingMillimeters,
                SpacingXMillimeters = input.SpacingXMillimeters,
                SpacingYMillimeters = input.SpacingYMillimeters,
                PileCapThicknessMillimeters = input.PileCapThicknessMillimeters,
                EdgeDistanceMillimeters = input.EdgeDistanceMillimeters,
                PileMaterial = input.PileMaterial,
                PileCapMaterial = input.PileCapMaterial
            };

            if (normalized.AutoSpacing)
            {
                double defaultSpacing = normalized.PileDiameterMillimeters * 3.0;
                normalized.PileSpacingMillimeters = defaultSpacing;
                normalized.SpacingXMillimeters = defaultSpacing;
                normalized.SpacingYMillimeters = defaultSpacing;
            }

            return normalized;
        }

        private PileCapGeometry CalculateMonoPileLayout(PileCapInputParameters input)
        {
            double capWidth = input.PileDiameterMillimeters + input.EdgeDistanceMillimeters * 2.0;
            double capLength = capWidth;
            List<PileCapPoint2D> boundary = CreateRectangleBoundary(capWidth, capLength);
            List<PileCapPoint2D> piles = new List<PileCapPoint2D>
            {
                new PileCapPoint2D(0, 0)
            };

            List<PileCapMeshArea> meshAreas = CreateRectangularGridMesh(
                CreateSortedUniqueValues(-capWidth / 2.0, 0, capWidth / 2.0),
                CreateSortedUniqueValues(-capLength / 2.0, 0, capLength / 2.0));

            return new PileCapGeometry(
                input.ArrangementType,
                piles,
                boundary,
                meshAreas,
                capWidth,
                capLength,
                0,
                0,
                input.RotationDegrees);
        }

        private PileCapGeometry CalculateTwoPileLayout(PileCapInputParameters input)
        {
            double spacing = input.PileSpacingMillimeters;
            double capWidth = spacing + input.PileDiameterMillimeters + input.EdgeDistanceMillimeters * 2.0;
            double capLength = input.PileDiameterMillimeters + input.EdgeDistanceMillimeters * 2.0;
            var piles = new List<PileCapPoint2D>
            {
                new PileCapPoint2D(-spacing / 2.0, 0),
                new PileCapPoint2D(spacing / 2.0, 0)
            };

            List<PileCapMeshArea> meshAreas = CreateRectangularGridMesh(
                CreateSortedUniqueValues(-capWidth / 2.0, -spacing / 2.0, 0, spacing / 2.0, capWidth / 2.0),
                CreateSortedUniqueValues(-capLength / 2.0, 0, capLength / 2.0));

            return new PileCapGeometry(
                input.ArrangementType,
                piles,
                CreateRectangleBoundary(capWidth, capLength),
                meshAreas,
                capWidth,
                capLength,
                spacing,
                0,
                input.RotationDegrees);
        }

        private PileCapGeometry CalculateThreePileLayout(PileCapInputParameters input)
        {
            double spacing = input.PileSpacingMillimeters;
            ThreePileCapGeometry threePileGeometry = new ThreePileCapBoundaryCalculator().Calculate(
                input.PileDiameterMillimeters,
                input.EdgeDistanceMillimeters,
                spacing);

            PileCapPoint2D bottomLeftPile = threePileGeometry.PileCenters[0];
            PileCapPoint2D bottomRightPile = threePileGeometry.PileCenters[1];
            PileCapPoint2D topPile = threePileGeometry.PileCenters[2];
            PileCapPoint2D center = new PileCapPoint2D(0, 0);
            IReadOnlyList<PileCapPoint2D> boundary = threePileGeometry.FinalBoundaryVertices;

            var meshAreas = new List<PileCapMeshArea>
            {
                new PileCapMeshArea(new[] { boundary[0], boundary[1], bottomRightPile, bottomLeftPile }),
                new PileCapMeshArea(new[] { boundary[1], boundary[2], bottomRightPile }),
                new PileCapMeshArea(new[] { boundary[2], boundary[3], topPile, bottomRightPile }),
                new PileCapMeshArea(new[] { boundary[3], boundary[4], topPile }),
                new PileCapMeshArea(new[] { boundary[4], boundary[5], bottomLeftPile, topPile }),
                new PileCapMeshArea(new[] { boundary[5], boundary[0], bottomLeftPile }),
                new PileCapMeshArea(new[] { bottomLeftPile, bottomRightPile, center }),
                new PileCapMeshArea(new[] { bottomRightPile, topPile, center }),
                new PileCapMeshArea(new[] { topPile, bottomLeftPile, center })
            };

            double minX;
            double maxX;
            double minY;
            double maxY;
            GetBounds(boundary, out minX, out maxX, out minY, out maxY);

            return new PileCapGeometry(
                input.ArrangementType,
                new List<PileCapPoint2D> { bottomLeftPile, bottomRightPile, topPile },
                boundary,
                meshAreas,
                maxX - minX,
                maxY - minY,
                spacing,
                0,
                input.RotationDegrees,
                threePileGeometry.SelectedPoint,
                threePileGeometry.TemporaryMonoCaps,
                threePileGeometry.FinalBoundarySegments);
        }

        private PileCapGeometry CalculateFourPileLayout(PileCapInputParameters input)
        {
            double spacingX = input.SpacingXMillimeters;
            double spacingY = input.SpacingYMillimeters;
            double capWidth = spacingX + input.PileDiameterMillimeters + input.EdgeDistanceMillimeters * 2.0;
            double capLength = spacingY + input.PileDiameterMillimeters + input.EdgeDistanceMillimeters * 2.0;
            var piles = new List<PileCapPoint2D>
            {
                new PileCapPoint2D(-spacingX / 2.0, -spacingY / 2.0),
                new PileCapPoint2D(spacingX / 2.0, -spacingY / 2.0),
                new PileCapPoint2D(spacingX / 2.0, spacingY / 2.0),
                new PileCapPoint2D(-spacingX / 2.0, spacingY / 2.0)
            };

            List<PileCapMeshArea> meshAreas = CreateRectangularGridMesh(
                CreateSortedUniqueValues(-capWidth / 2.0, -spacingX / 2.0, 0, spacingX / 2.0, capWidth / 2.0),
                CreateSortedUniqueValues(-capLength / 2.0, -spacingY / 2.0, 0, spacingY / 2.0, capLength / 2.0));

            return new PileCapGeometry(
                input.ArrangementType,
                piles,
                CreateRectangleBoundary(capWidth, capLength),
                meshAreas,
                capWidth,
                capLength,
                spacingX,
                spacingY,
                input.RotationDegrees);
        }

        private static List<PileCapPoint2D> CreateRectangleBoundary(double width, double length)
        {
            double halfWidth = width / 2.0;
            double halfLength = length / 2.0;
            return new List<PileCapPoint2D>
            {
                new PileCapPoint2D(-halfWidth, -halfLength),
                new PileCapPoint2D(halfWidth, -halfLength),
                new PileCapPoint2D(halfWidth, halfLength),
                new PileCapPoint2D(-halfWidth, halfLength)
            };
        }

        private static List<PileCapMeshArea> CreateRectangularGridMesh(IReadOnlyList<double> xs, IReadOnlyList<double> ys)
        {
            var meshAreas = new List<PileCapMeshArea>();
            for (int yi = 0; yi < ys.Count - 1; yi++)
            {
                for (int xi = 0; xi < xs.Count - 1; xi++)
                {
                    meshAreas.Add(new PileCapMeshArea(new[]
                    {
                        new PileCapPoint2D(xs[xi], ys[yi]),
                        new PileCapPoint2D(xs[xi + 1], ys[yi]),
                        new PileCapPoint2D(xs[xi + 1], ys[yi + 1]),
                        new PileCapPoint2D(xs[xi], ys[yi + 1])
                    }));
                }
            }

            return meshAreas;
        }

        private static List<double> CreateSortedUniqueValues(params double[] values)
        {
            var sorted = new List<double>(values ?? new double[0]);
            sorted.Sort();

            var unique = new List<double>();
            foreach (double value in sorted)
            {
                if (unique.Count == 0 || Math.Abs(unique[unique.Count - 1] - value) > Tolerance)
                {
                    unique.Add(value);
                }
            }

            return unique;
        }

        private static void GetBounds(
            IReadOnlyList<PileCapPoint2D> points,
            out double minX,
            out double maxX,
            out double minY,
            out double maxY)
        {
            minX = 0;
            maxX = 0;
            minY = 0;
            maxY = 0;
            if (points == null || points.Count == 0)
            {
                return;
            }

            minX = maxX = points[0].X;
            minY = maxY = points[0].Y;
            for (int i = 1; i < points.Count; i++)
            {
                minX = Math.Min(minX, points[i].X);
                maxX = Math.Max(maxX, points[i].X);
                minY = Math.Min(minY, points[i].Y);
                maxY = Math.Max(maxY, points[i].Y);
            }
        }

    }
}
