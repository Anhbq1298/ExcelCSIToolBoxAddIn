using System.Collections.Generic;
using ExcelCSIToolBox.Core.Models.CSI;
using ExcelCSIToolBox.Core.Tabular;
using FluentAssertions;
using Xunit;

namespace ExcelCSIToolBox.Tests
{
    public class CsiTableFieldAliasResolverTests
    {
        [Fact]
        public void FindObjectNameColumn_resolves_frame_specific_aliases()
        {
            var fields = new List<string> { "Story", "Beam", "Output Case" };

            int index = CsiTableFieldAliasResolver.FindObjectNameColumn(fields, CsiObjectTypes.Frame);

            index.Should().Be(1);
        }

        [Fact]
        public void FindObjectNameColumn_resolves_area_shell_aliases()
        {
            var fields = new List<string> { "Story", "Shell Name", "F11" };

            int index = CsiTableFieldAliasResolver.FindObjectNameColumn(fields, CsiObjectTypes.Area);

            index.Should().Be(1);
        }

        [Fact]
        public void FindObjectNameColumn_resolves_point_joint_aliases_case_insensitively()
        {
            var fields = new List<string> { "story", "jointname", "u1" };

            int index = CsiTableFieldAliasResolver.FindObjectNameColumn(fields, CsiObjectTypes.Point);

            index.Should().Be(1);
        }
    }
}
