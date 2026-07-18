using System.Globalization;
using System.Threading;
using ExcelCSIToolBox.Application.Modelling.DropPanels;
using FluentAssertions;
using Xunit;

namespace ExcelCSIToolBox.Tests.Application.Modelling.DropPanels
{
    public class DropPanelPropertyNameBuilderTests
    {
        [Theory]
        [InlineData(1500.0, "C32/40", "Drop_1500_C32/40")]
        [InlineData(250.0, "C40/50", "Drop_250_C40/50")]
        [InlineData(275.5, "C32/40", "Drop_275.5_C32/40")]
        public void Build_UsesDeterministicThicknessAndExactMaterialName(
            double thickness,
            string material,
            string expectedName)
        {
            var result = DropPanelPropertyNameBuilder.Build(thickness, material);

            result.IsSuccess.Should().BeTrue();
            result.Data.Should().Be(expectedName);
        }

        [Fact]
        public void Build_UsesDotDecimalSeparatorRegardlessOfCurrentCulture()
        {
            CultureInfo originalCulture = Thread.CurrentThread.CurrentCulture;
            try
            {
                Thread.CurrentThread.CurrentCulture = new CultureInfo("vi-VN");

                var result = DropPanelPropertyNameBuilder.Build(275.5, "C32/40");

                result.IsSuccess.Should().BeTrue();
                result.Data.Should().Be("Drop_275.5_C32/40");
            }
            finally
            {
                Thread.CurrentThread.CurrentCulture = originalCulture;
            }
        }

        [Theory]
        [InlineData(0.0)]
        [InlineData(-1.0)]
        [InlineData(double.NaN)]
        [InlineData(double.PositiveInfinity)]
        public void Build_RejectsInvalidThickness(double thickness)
        {
            var result = DropPanelPropertyNameBuilder.Build(thickness, "C32/40");

            result.IsSuccess.Should().BeFalse();
            result.Message.Should().Contain("greater than zero");
        }

        [Fact]
        public void Build_RejectsMissingMaterial()
        {
            var result = DropPanelPropertyNameBuilder.Build(1500.0, "  ");

            result.IsSuccess.Should().BeFalse();
            result.Message.Should().Contain("concrete material");
        }
    }
}
