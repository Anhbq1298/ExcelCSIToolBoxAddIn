using System.Collections.Generic;
using ExcelCSIToolBox.Application.Models.Export;
using ExcelCSIToolBox.Application.Services;
using ExcelCSIToolBox.Core.Common.Results;
using ExcelCSIToolBox.Core.Models.CSI;
using ExcelCSIToolBox.Core.Models.EtabsTables;
using FluentAssertions;
using Xunit;

namespace ExcelCSIToolBox.Tests
{
    public class SelectedObjectTableFilterTests
    {
        [Fact]
        public void Filter_matches_unique_name()
        {
            var table = new EtabsTableResult
            {
                TableName = "Beam Object Connectivity",
                Headers = new List<string> { "UniqueName", "Story" },
                Rows = new List<List<string>>
                {
                    new List<string> { "B1", "L1" },
                    new List<string> { "B2", "L1" }
                }
            };
            var identities = new[]
            {
                CsiObjectIdentity.Create(CsiObjectTypes.Frame, "B2", "Beam 2", "L1")
            };

            OperationResult<PreparedTableExport> result = SelectedObjectTableFilter.Filter(
                table,
                identities,
                CsiObjectTypes.Frame,
                "Beam Object Connectivity");

            result.IsSuccess.Should().BeTrue();
            result.Data.RecordCount.Should().Be(1);
            result.Data.Rows[0][0].Should().Be("B2");
        }

        [Fact]
        public void Filter_matches_label_case_insensitively()
        {
            var table = new EtabsTableResult
            {
                TableName = "Element Forces - Beams",
                Headers = new List<string> { "Story", "Beam", "P" },
                Rows = new List<List<string>>
                {
                    new List<string> { "L1", "B-A", "10" },
                    new List<string> { "L1", "B-B", "20" }
                }
            };
            var identities = new[]
            {
                CsiObjectIdentity.Create(CsiObjectTypes.Frame, "1001", "b-a", "L1")
            };

            OperationResult<PreparedTableExport> result = SelectedObjectTableFilter.Filter(
                table,
                identities,
                CsiObjectTypes.Frame,
                "Element Forces - Beams");

            result.IsSuccess.Should().BeTrue();
            result.Data.RecordCount.Should().Be(1);
            result.Data.Rows[0][1].Should().Be("B-A");
        }

        [Fact]
        public void Filter_returns_failure_for_empty_selection()
        {
            var table = new EtabsTableResult
            {
                TableName = "Point Object Connectivity",
                Headers = new List<string> { "UniqueName" },
                Rows = new List<List<string>> { new List<string> { "1" } }
            };

            OperationResult<PreparedTableExport> result = SelectedObjectTableFilter.Filter(
                table,
                new CsiObjectIdentity[0],
                CsiObjectTypes.Point,
                "Point Object Connectivity");

            result.IsSuccess.Should().BeFalse();
            result.Message.Should().Contain("Select one or more joint objects");
        }

        [Fact]
        public void Filter_returns_failure_when_object_column_is_missing()
        {
            var table = new EtabsTableResult
            {
                TableName = "Unknown",
                Headers = new List<string> { "Story", "Length" },
                Rows = new List<List<string>> { new List<string> { "L1", "1.0" } }
            };
            var identities = new[]
            {
                CsiObjectIdentity.Create(CsiObjectTypes.Frame, "B1", "B1", "L1")
            };

            OperationResult<PreparedTableExport> result = SelectedObjectTableFilter.Filter(
                table,
                identities,
                CsiObjectTypes.Frame,
                "Unknown");

            result.IsSuccess.Should().BeFalse();
            result.Message.Should().Contain("object name field");
        }
    }
}
