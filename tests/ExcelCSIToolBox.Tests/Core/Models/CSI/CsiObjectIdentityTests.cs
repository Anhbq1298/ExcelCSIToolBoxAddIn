using System.Linq;
using ExcelCSIToolBox.Core.Models.CSI;
using FluentAssertions;
using Xunit;

namespace ExcelCSIToolBox.Tests
{
    public class CsiObjectIdentityTests
    {
        [Theory]
        [InlineData(CsiObjectTypes.Point)]
        [InlineData(CsiObjectTypes.Frame)]
        [InlineData(CsiObjectTypes.Area)]
        public void Create_includes_unique_name_label_and_story_match_keys(string objectType)
        {
            CsiObjectIdentity identity = CsiObjectIdentity.Create(objectType, " 101 ", " C1 ", " STORY1 ");

            identity.ObjectType.Should().Be(objectType);
            identity.UniqueName.Should().Be("101");
            identity.Label.Should().Be("C1");
            identity.Story.Should().Be("STORY1");
            identity.MatchKeys.Should().Contain(new[] { "101", "C1", "STORY1" });
            identity.Matches("c1").Should().BeTrue();
        }

        [Fact]
        public void Create_falls_back_to_unique_name_when_label_is_missing()
        {
            CsiObjectIdentity identity = CsiObjectIdentity.Create(CsiObjectTypes.Frame, "B-1", null, null);

            identity.MatchKeys.Should().ContainSingle().Which.Should().Be("B-1");
            identity.Matches("b-1").Should().BeTrue();
        }

        [Fact]
        public void Create_normalizes_invalid_object_type_to_unknown_but_keeps_match_key()
        {
            CsiObjectIdentity identity = CsiObjectIdentity.Create("999", "raw-name", null, null);

            identity.ObjectType.Should().Be(CsiObjectTypes.Unknown);
            identity.MatchKeys.Single().Should().Be("raw-name");
        }
    }
}
