using ExcelCSIToolBox.Application.Services;
using ExcelCSIToolBox.Core.Models.CSI;
using ExcelCSIToolBox.Core.Models.ElementConnectivity;
using FluentAssertions;
using Xunit;

namespace ExcelCSIToolBox.Tests
{
    public class ObjectConnectivityRequestFactoryTests
    {
        [Theory]
        [InlineData("POINT_OBJECT_CONNECTIVITY", "Point Object Connectivity", CsiObjectTypes.Point)]
        [InlineData("BEAM_OBJECT_CONNECTIVITY", "Beam Object Connectivity", CsiObjectTypes.Frame)]
        [InlineData("COLUMN_OBJECT_CONNECTIVITY", "Column Object Connectivity", CsiObjectTypes.Frame)]
        [InlineData("BRACE_OBJECT_CONNECTIVITY", "Brace Object Connectivity", CsiObjectTypes.Frame)]
        [InlineData("FLOOR_OBJECT_CONNECTIVITY", "Floor Object Connectivity", CsiObjectTypes.Area)]
        [InlineData("WALL_OBJECT_CONNECTIVITY", "Wall Object Connectivity", CsiObjectTypes.Area)]
        [InlineData("OTHER", "Other Connectivity", CsiObjectTypes.Unknown)]
        public void Create_ResolvesObjectCategoryFromConnectivityItem(string key, string title, string expectedCategory)
        {
            var item = new ElementConnectivityItem(title, key, "Group", "Table Name");

            var request = ObjectConnectivityRequestFactory.Create(item);

            request.TableName.Should().Be("Table Name");
            request.DisplayName.Should().Be(title);
            request.ObjectCategory.Should().Be(expectedCategory);
        }

        [Fact]
        public void Create_ReturnsEmptyUnknownRequest_WhenItemIsNull()
        {
            var request = ObjectConnectivityRequestFactory.Create(null);

            request.TableName.Should().BeEmpty();
            request.DisplayName.Should().BeEmpty();
            request.ObjectCategory.Should().Be(CsiObjectTypes.Unknown);
        }
    }
}
