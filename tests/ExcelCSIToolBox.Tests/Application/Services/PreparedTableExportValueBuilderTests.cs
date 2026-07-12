using System.Collections.Generic;
using ExcelCSIToolBox.Application.Models.Export;
using ExcelCSIToolBox.Application.Services;
using FluentAssertions;
using Xunit;

namespace ExcelCSIToolBox.Tests
{
    public class PreparedTableExportValueBuilderTests
    {
        [Fact]
        public void BuildValues_IncludesHeadersAndPadsShortRows()
        {
            var export = new PreparedTableExport
            {
                Headers = new[] { "Name", "Type", "Story" },
                Rows = new List<IReadOnlyList<object>>
                {
                    new object[] { "B1", "Beam", "L2" },
                    new object[] { "C1" }
                }
            };

            object[,] values = PreparedTableExportValueBuilder.BuildValues(export, true);

            values.GetLength(0).Should().Be(3);
            values.GetLength(1).Should().Be(3);
            values[0, 0].Should().Be("Name");
            values[0, 1].Should().Be("Type");
            values[1, 2].Should().Be("L2");
            values[2, 0].Should().Be("C1");
            values[2, 1].Should().Be(string.Empty);
            values[2, 2].Should().Be(string.Empty);
        }

        [Fact]
        public void BuildValues_UsesLongestDataRow_WhenHeadersAreMissing()
        {
            var export = new PreparedTableExport
            {
                Rows = new List<IReadOnlyList<object>>
                {
                    new object[] { "A", "B", "C" }
                }
            };

            object[,] values = PreparedTableExportValueBuilder.BuildValues(export, false);

            values.GetLength(0).Should().Be(1);
            values.GetLength(1).Should().Be(3);
            values[0, 2].Should().Be("C");
        }

        [Fact]
        public void BuildValues_ReturnsEmptyMatrix_WhenNoColumnsAreAvailable()
        {
            object[,] values = PreparedTableExportValueBuilder.BuildValues(new PreparedTableExport(), true);

            values.GetLength(0).Should().Be(0);
            values.GetLength(1).Should().Be(0);
        }
    }
}
