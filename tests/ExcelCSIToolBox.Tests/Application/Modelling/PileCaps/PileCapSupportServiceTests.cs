using System.Linq;
using ExcelCSIToolBox.Application.Modelling.PileCaps;
using ExcelCSIToolBox.Core.Contracts.CSI.PileCap;
using FluentAssertions;
using Xunit;

namespace ExcelCSIToolBox.Tests.Application.Modelling.PileCaps
{
    public class PileCapSupportServiceTests
    {
        [Fact]
        public void BuildPileFrameSectionName_UsesIntegerMillimetersAndRawMaterialName()
        {
            string name = PileCapPropertyNameBuilder.BuildPileFrameSectionName(800.4, "C32/40");

            name.Should().Be("P_800D_C32/40");
        }

        [Fact]
        public void BuildPileCapAreaSectionName_UsesIntegerMillimetersAndRawMaterialName()
        {
            string name = PileCapPropertyNameBuilder.BuildPileCapAreaSectionName(1500, "C32/40");

            name.Should().Be("PC_1500_C32/40");
        }

        [Fact]
        public void EtabsUnitConverter_ConvertsMetersAndMillimeters()
        {
            var converter = new EtabsUnitConverter();

            converter.MillimetersToModelLength(30000, 6).Should().BeApproximately(30, 0.0001);
            converter.ModelLengthToMillimeters(1.5, 6).Should().BeApproximately(1500, 0.0001);
            converter.MillimetersToModelLength(800, 4).Should().BeApproximately(800, 0.0001);
        }

        [Fact]
        public void Validate_FailsWhenSpacingIsNotGreaterThanPileDiameter()
        {
            var validator = new PileCapInputValidator();
            var input = new PileCapInputParameters
            {
                ArrangementType = PileCapArrangementType.TwoPile,
                PileDiameterMillimeters = 800,
                PileLengthMillimeters = 30000,
                PileSpacingMillimeters = 800,
                PileCapThicknessMillimeters = 1500,
                EdgeDistanceMillimeters = 150,
                PileMaterial = "C32/40",
                PileCapMaterial = "C32/40"
            };

            validator.Validate(input).Should().Contain(message => message.Contains("Pile spacing"));
        }

        [Fact]
        public void Validate_FailsWhenFourPileSpacingYIsNotGreaterThanPileDiameter()
        {
            var validator = new PileCapInputValidator();
            var input = new PileCapInputParameters
            {
                ArrangementType = PileCapArrangementType.FourPile,
                PileDiameterMillimeters = 800,
                PileLengthMillimeters = 30000,
                SpacingXMillimeters = 2400,
                SpacingYMillimeters = 700,
                PileCapThicknessMillimeters = 1500,
                EdgeDistanceMillimeters = 150,
                PileMaterial = "C32/40",
                PileCapMaterial = "C32/40"
            };

            validator.Validate(input).Any(message => message.Contains("Spacing Y")).Should().BeTrue();
        }
    }
}
