using System;
using ExcelCSIToolBox.Infrastructure.CSI.Common;
using FluentAssertions;
using Xunit;

namespace ExcelCSIToolBox.Tests
{
    public class CurrentThreadCsiApiDispatcherTests
    {
        [Fact]
        public void Invoke_runs_action()
        {
            var dispatcher = new CurrentThreadCsiApiDispatcher();
            bool called = false;

            dispatcher.Invoke(() => called = true);

            called.Should().BeTrue();
        }

        [Fact]
        public void Invoke_returns_result()
        {
            var dispatcher = new CurrentThreadCsiApiDispatcher();

            int result = dispatcher.Invoke(() => 42);

            result.Should().Be(42);
        }

        [Fact]
        public void Invoke_propagates_exception()
        {
            var dispatcher = new CurrentThreadCsiApiDispatcher();

            Action action = () => dispatcher.Invoke(() => { throw new InvalidOperationException("bad"); });

            action.Should().Throw<InvalidOperationException>().WithMessage("bad");
        }
    }
}
