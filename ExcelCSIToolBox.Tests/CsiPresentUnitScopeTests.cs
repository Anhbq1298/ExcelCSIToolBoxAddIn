using ExcelCSIToolBox.Application.Services;
using ExcelCSIToolBox.Core.Abstractions.CSI;
using ExcelCSIToolBox.Core.Common.Results;
using ExcelCSIToolBox.Data.DTOs.CSI;
using FluentAssertions;
using NSubstitute;
using Xunit;

namespace ExcelCSIToolBox.Tests
{
    public class CsiPresentUnitScopeTests
    {
        [Fact]
        public void Apply_SetsRequestedUnits_AndRestoresOriginalUnitsOnDispose()
        {
            var connection = Substitute.For<ICSISapModelConnectionService>();
            var original = Unit(4, 6, 2);
            var requested = Unit(3, 4, 2);
            connection.GetPresentUnitSystem().Returns(OperationResult<CSISapModelPresentUnitSystemDTO>.Success(original));
            connection.SetPresentUnitSystem(Arg.Any<CSISapModelPresentUnitSystemDTO>()).Returns(OperationResult.Success());

            OperationResult<CsiPresentUnitScope> result = CsiPresentUnitScope.Apply(connection, requested);
            result.IsSuccess.Should().BeTrue();

            connection.Received(1).SetPresentUnitSystem(Arg.Is<CSISapModelPresentUnitSystemDTO>(u =>
                u.ForceUnit == requested.ForceUnit &&
                u.LengthUnit == requested.LengthUnit &&
                u.TemperatureUnit == requested.TemperatureUnit));

            result.Data.Dispose();

            connection.Received(1).SetPresentUnitSystem(Arg.Is<CSISapModelPresentUnitSystemDTO>(u =>
                u.ForceUnit == original.ForceUnit &&
                u.LengthUnit == original.LengthUnit &&
                u.TemperatureUnit == original.TemperatureUnit));
            result.Data.RestoreResult.IsSuccess.Should().BeTrue();
        }

        [Fact]
        public void Apply_FailsWithoutChangingUnits_WhenCurrentUnitsCannotBeRead()
        {
            var connection = Substitute.For<ICSISapModelConnectionService>();
            connection.GetPresentUnitSystem().Returns(OperationResult<CSISapModelPresentUnitSystemDTO>.Failure("read failed"));

            OperationResult<CsiPresentUnitScope> result = CsiPresentUnitScope.Apply(connection, Unit(3, 4, 2));

            result.IsSuccess.Should().BeFalse();
            result.Message.Should().Be("read failed");
            connection.DidNotReceive().SetPresentUnitSystem(Arg.Any<CSISapModelPresentUnitSystemDTO>());
        }

        [Fact]
        public void Apply_Fails_WhenRequestedUnitsCannotBeApplied()
        {
            var connection = Substitute.For<ICSISapModelConnectionService>();
            connection.GetPresentUnitSystem().Returns(OperationResult<CSISapModelPresentUnitSystemDTO>.Success(Unit(4, 6, 2)));
            connection.SetPresentUnitSystem(Arg.Any<CSISapModelPresentUnitSystemDTO>())
                .Returns(OperationResult.Failure("apply failed"));

            OperationResult<CsiPresentUnitScope> result = CsiPresentUnitScope.Apply(connection, Unit(3, 4, 2));

            result.IsSuccess.Should().BeFalse();
            result.Message.Should().Be("apply failed");
        }

        [Fact]
        public void Dispose_CapturesRestoreFailureWithoutThrowing()
        {
            var connection = Substitute.For<ICSISapModelConnectionService>();
            int setCallCount = 0;
            connection.GetPresentUnitSystem().Returns(OperationResult<CSISapModelPresentUnitSystemDTO>.Success(Unit(4, 6, 2)));
            connection.SetPresentUnitSystem(Arg.Any<CSISapModelPresentUnitSystemDTO>())
                .Returns(_ =>
                {
                    setCallCount++;
                    return setCallCount == 1
                        ? OperationResult.Success()
                        : OperationResult.Failure("restore failed");
                });

            OperationResult<CsiPresentUnitScope> result = CsiPresentUnitScope.Apply(connection, Unit(3, 4, 2));

            result.Data.Dispose();

            result.Data.RestoreResult.IsSuccess.Should().BeFalse();
            result.Data.RestoreResult.Message.Should().Be("restore failed");
        }

        private static CSISapModelPresentUnitSystemDTO Unit(int force, int length, int temperature)
        {
            return new CSISapModelPresentUnitSystemDTO
            {
                ForceUnit = force,
                LengthUnit = length,
                TemperatureUnit = temperature
            };
        }
    }
}
