using ExcelCSIToolBox.Core.Common.Results;
using FluentAssertions;
using Xunit;

namespace ExcelCSIToolBox.Tests
{
    public class OperationResultTests
    {
        [Fact]
        public void Success_Sets_success_flag_and_message()
        {
            OperationResult result = OperationResult.Success("done");

            result.IsSuccess.Should().BeTrue();
            result.Message.Should().Be("done");
        }

        [Fact]
        public void Failure_Sets_failure_flag_and_message()
        {
            OperationResult result = OperationResult.Failure("bad");

            result.IsSuccess.Should().BeFalse();
            result.Message.Should().Be("bad");
        }

        [Fact]
        public void Generic_success_carries_data()
        {
            OperationResult<int> result = OperationResult<int>.Success(42, "ok");

            result.IsSuccess.Should().BeTrue();
            result.Data.Should().Be(42);
            result.Message.Should().Be("ok");
        }

        [Fact]
        public void Generic_failure_preserves_error_and_default_data()
        {
            OperationResult<string> result = OperationResult<string>.Failure("missing");

            result.IsSuccess.Should().BeFalse();
            result.Message.Should().Be("missing");
            result.Data.Should().BeNull();
        }
    }
}
