using System;
using System.Linq;
using ExcelCSIToolBox.Application.Modelling.PileCaps;
using ExcelCSIToolBox.Core.Contracts.CSI.PileCap;
using FluentAssertions;
using Xunit;

namespace ExcelCSIToolBox.Tests.Application.Modelling.PileCaps
{
    public class PileCapGeometryCalculatorTests
    {
        private readonly PileCapGeometryCalculator _calculator = new PileCapGeometryCalculator();

        [Fact]
        public void Calculate_MonoPileCap_CreatesCenteredSquareCap()
        {
            PileCapGeometry geometry = _calculator.Calculate(CreateInput(PileCapArrangementType.Mono));

            geometry.PileCenters.Should().ContainSingle();
            geometry.CapWidthXMillimeters.Should().BeApproximately(1100, 0.0001);
            geometry.CapLengthYMillimeters.Should().BeApproximately(1100, 0.0001);
            geometry.BoundaryVertices.Should().HaveCount(4);
            geometry.MeshAreas.Should().HaveCount(4);
        }

        [Fact]
        public void Calculate_TwoPileCap_CreatesHorizontalGroupAndExpectedCap()
        {
            PileCapGeometry geometry = _calculator.Calculate(CreateInput(PileCapArrangementType.TwoPile));

            geometry.PileCenters.Should().HaveCount(2);
            geometry.PileCenters[0].X.Should().BeApproximately(-1200, 0.0001);
            geometry.PileCenters[1].X.Should().BeApproximately(1200, 0.0001);
            geometry.PileCenters[0].Y.Should().BeApproximately(0, 0.0001);
            geometry.PileCenters[1].Y.Should().BeApproximately(0, 0.0001);
            geometry.CapWidthXMillimeters.Should().BeApproximately(3500, 0.0001);
            geometry.CapLengthYMillimeters.Should().BeApproximately(1100, 0.0001);
        }

        [Fact]
        public void Calculate_ThreePileCap_CreatesSixVertexBoundaryFromTemporaryMonoCaps()
        {
            PileCapGeometry geometry = _calculator.Calculate(CreateInput(PileCapArrangementType.ThreePile));

            geometry.PileCenters.Should().HaveCount(3);
            geometry.BoundaryVertices.Should().HaveCount(6);
            geometry.BoundarySegments.Should().HaveCount(6);
            geometry.MeshAreas.Should().HaveCount(9);
            geometry.TemporaryMonoCaps.Should().HaveCount(3);

            AssertPoint(geometry.PileCenters[0], -1200, -692.820323);
            AssertPoint(geometry.PileCenters[1], 1200, -692.820323);
            AssertPoint(geometry.PileCenters[2], 0, 1385.640646);

            AssertPoint(geometry.BoundaryVertices[0], -1750, -1242.820323);
            AssertPoint(geometry.BoundaryVertices[1], 1750, -1242.820323);
            AssertPoint(geometry.BoundaryVertices[2], 1750, -142.820323);
            AssertPoint(geometry.BoundaryVertices[3], 550, 1935.640646);
            AssertPoint(geometry.BoundaryVertices[4], -550, 1935.640646);
            AssertPoint(geometry.BoundaryVertices[5], -1750, -142.820323);

            double centroidX = geometry.PileCenters.Average(point => point.X);
            double centroidY = geometry.PileCenters.Average(point => point.Y);
            centroidX.Should().BeApproximately(0, 0.0001);
            centroidY.Should().BeApproximately(0, 0.0001);
            AssertPoint(geometry.SelectedPoint, 0, 0);

            geometry.BoundaryVertices[0].Y.Should().BeApproximately(geometry.BoundaryVertices[1].Y, 0.0001);
            geometry.BoundaryVertices[3].Y.Should().BeApproximately(geometry.BoundaryVertices[4].Y, 0.0001);
            geometry.BoundaryVertices[1].X.Should().BeApproximately(geometry.BoundaryVertices[2].X, 0.0001);
            geometry.BoundaryVertices[5].X.Should().BeApproximately(geometry.BoundaryVertices[0].X, 0.0001);
            geometry.CapWidthXMillimeters.Should().BeApproximately(3500, 0.0001);
            geometry.CapLengthYMillimeters.Should().BeApproximately(3178.460969, 0.0001);
            geometry.SpacingXMillimeters.Should().BeApproximately(2400, 0.0001);
            geometry.SpacingYMillimeters.Should().BeApproximately(0, 0.0001);

            Distance(geometry.PileCenters[0], geometry.PileCenters[1]).Should().BeApproximately(2400, 0.0001);
            Distance(geometry.PileCenters[1], geometry.PileCenters[2]).Should().BeApproximately(2400, 0.0001);
            Distance(geometry.PileCenters[2], geometry.PileCenters[0]).Should().BeApproximately(2400, 0.0001);

            double radius = Distance(geometry.SelectedPoint, geometry.PileCenters[0]);
            Distance(geometry.SelectedPoint, geometry.PileCenters[1]).Should().BeApproximately(radius, 0.0001);
            Distance(geometry.SelectedPoint, geometry.PileCenters[2]).Should().BeApproximately(radius, 0.0001);

            foreach (Rectangle2D temporaryMonoCap in geometry.TemporaryMonoCaps)
            {
                temporaryMonoCap.Width.Should().BeApproximately(1100, 0.0001);
                temporaryMonoCap.Height.Should().BeApproximately(1100, 0.0001);
                foreach (PileCapPoint2D corner in temporaryMonoCap.GetCornersClockwise())
                {
                    ThreePileCapBoundaryCalculator
                        .IsPointInsideOrOnBoundary(corner, geometry.BoundaryVertices)
                        .Should()
                        .BeTrue();
                }
            }

            geometry.BoundarySegments.Select(segment => segment.Start.X + "," + segment.Start.Y + "->" + segment.End.X + "," + segment.End.Y)
                .Distinct()
                .Should()
                .HaveCount(6);
        }

        [Fact]
        public void Calculate_FourPileCap_CreatesTwoByTwoGroupAndExpectedCap()
        {
            PileCapGeometry geometry = _calculator.Calculate(CreateInput(PileCapArrangementType.FourPile));

            geometry.PileCenters.Should().HaveCount(4);
            geometry.CapWidthXMillimeters.Should().BeApproximately(3500, 0.0001);
            geometry.CapLengthYMillimeters.Should().BeApproximately(3500, 0.0001);
            geometry.MeshAreas.Should().HaveCount(16);
        }

        [Fact]
        public void RotatePoints_RotatesAroundSelectedPoint()
        {
            PileCapGeometry geometry = _calculator.Calculate(new PileCapInputParameters
            {
                ArrangementType = PileCapArrangementType.TwoPile,
                PileDiameterMillimeters = 800,
                PileLengthMillimeters = 30000,
                PileSpacingMillimeters = 2400,
                PileCapThicknessMillimeters = 1500,
                EdgeDistanceMillimeters = 150,
                RotationDegrees = 90,
                PileMaterial = "C32/40",
                PileCapMaterial = "C32/40"
            });

            var rotated = _calculator.RotatePoints(geometry.PileCenters, 90);

            rotated[0].X.Should().BeApproximately(0, 0.0001);
            rotated[0].Y.Should().BeApproximately(-1200, 0.0001);
            rotated[1].X.Should().BeApproximately(0, 0.0001);
            rotated[1].Y.Should().BeApproximately(1200, 0.0001);
        }

        [Fact]
        public void Calculate_FourPileCap_AllowsDifferentSpacingXAndSpacingY()
        {
            PileCapGeometry geometry = _calculator.Calculate(new PileCapInputParameters
            {
                ArrangementType = PileCapArrangementType.FourPile,
                PileDiameterMillimeters = 800,
                PileLengthMillimeters = 30000,
                SpacingXMillimeters = 3000,
                SpacingYMillimeters = 2400,
                PileCapThicknessMillimeters = 1500,
                EdgeDistanceMillimeters = 150,
                PileMaterial = "C32/40",
                PileCapMaterial = "C32/40"
            });

            geometry.CapWidthXMillimeters.Should().BeApproximately(4100, 0.0001);
            geometry.CapLengthYMillimeters.Should().BeApproximately(3500, 0.0001);
        }

        [Fact]
        public void RotatePoints_ThreePileCap_PreservesDistancesAndBoundaryClosure()
        {
            PileCapGeometry geometry = _calculator.Calculate(new PileCapInputParameters
            {
                ArrangementType = PileCapArrangementType.ThreePile,
                PileDiameterMillimeters = 800,
                PileLengthMillimeters = 30000,
                PileSpacingMillimeters = 2400,
                PileCapThicknessMillimeters = 1500,
                EdgeDistanceMillimeters = 150,
                RotationDegrees = 30,
                PileMaterial = "C32/40",
                PileCapMaterial = "C32/40"
            });

            var rotatedPiles = _calculator.RotatePoints(geometry.PileCenters, geometry.RotationDegrees);
            var rotatedBoundary = _calculator.RotatePoints(geometry.BoundaryVertices, geometry.RotationDegrees);

            rotatedBoundary.Should().HaveCount(6);
            Distance(rotatedPiles[0], rotatedPiles[1]).Should().BeApproximately(2400, 0.0001);
            Distance(rotatedPiles[1], rotatedPiles[2]).Should().BeApproximately(2400, 0.0001);
            Distance(rotatedPiles[2], rotatedPiles[0]).Should().BeApproximately(2400, 0.0001);
            Distance(rotatedBoundary[5], rotatedBoundary[0])
                .Should()
                .BeApproximately(Distance(geometry.BoundaryVertices[5], geometry.BoundaryVertices[0]), 0.0001);
        }

        private static PileCapInputParameters CreateInput(PileCapArrangementType arrangementType)
        {
            return new PileCapInputParameters
            {
                ArrangementType = arrangementType,
                PileDiameterMillimeters = 800,
                PileLengthMillimeters = 30000,
                PileSpacingMillimeters = 2400,
                SpacingXMillimeters = 2400,
                SpacingYMillimeters = 2400,
                PileCapThicknessMillimeters = 1500,
                EdgeDistanceMillimeters = 150,
                PileMaterial = "C32/40",
                PileCapMaterial = "C32/40"
            };
        }

        private static double Distance(PileCapPoint2D left, PileCapPoint2D right)
        {
            double dx = left.X - right.X;
            double dy = left.Y - right.Y;
            return Math.Sqrt(dx * dx + dy * dy);
        }

        private static void AssertPoint(PileCapPoint2D point, double expectedX, double expectedY)
        {
            point.X.Should().BeApproximately(expectedX, 0.0001);
            point.Y.Should().BeApproximately(expectedY, 0.0001);
        }
    }
}
