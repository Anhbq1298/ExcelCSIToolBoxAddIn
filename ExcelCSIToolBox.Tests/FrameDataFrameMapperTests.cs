using System.Collections.Generic;
using ExcelCSIToolBox.Application.Mappers;
using FluentAssertions;
using Xunit;

namespace ExcelCSIToolBox.Tests
{
    public class FrameDataFrameMapperTests
    {
        [Fact]
        public void Map_creates_unique_name_column()
        {
            var dataFrame = CSISapModelFrameDataDataFrameMapper.Map(new[] { "F1" });

            dataFrame.Columns.Should().Equal("UniqueName");
        }

        [Fact]
        public void Map_creates_one_row_per_frame_name()
        {
            var dataFrame = CSISapModelFrameDataDataFrameMapper.Map(new[] { "F1", "F2" });

            dataFrame.Rows.Should().HaveCount(2);
            dataFrame.Rows[0].Should().Equal(new List<object> { "F1" });
            dataFrame.Rows[1].Should().Equal(new List<object> { "F2" });
        }
    }
}
