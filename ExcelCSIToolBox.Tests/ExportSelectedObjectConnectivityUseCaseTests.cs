using System.Collections.Generic;
using System.Threading.Tasks;
using ExcelCSIToolBox.Application.Interfaces.Etabs;
using ExcelCSIToolBox.Application.Models.Export;
using ExcelCSIToolBox.Application.UseCases;
using ExcelCSIToolBox.Core.Abstractions.CSI;
using ExcelCSIToolBox.Core.Common.Results;
using ExcelCSIToolBox.Core.Models.CSI;
using ExcelCSIToolBox.Core.Models.EtabsTables;
using FluentAssertions;
using NSubstitute;
using Xunit;

namespace ExcelCSIToolBox.Tests
{
    public class ExportSelectedObjectConnectivityUseCaseTests
    {
        [Fact]
        public async Task ExecuteAsync_filters_rows_using_resolved_selected_identities()
        {
            var tableService = Substitute.For<IEtabsDatabaseTableService>();
            tableService.GetTableAsync("Beam Object Connectivity").Returns(Task.FromResult(new EtabsTableResult
            {
                TableName = "Beam Object Connectivity",
                Headers = new List<string> { "UniqueName", "Story" },
                Rows = new List<List<string>>
                {
                    new List<string> { "B1", "L1" },
                    new List<string> { "B2", "L1" }
                }
            }));

            var identityResolver = Substitute.For<ISelectedObjectIdentityResolver>();
            identityResolver.ResolveSelectedObjects().Returns(
                OperationResult<IReadOnlyList<CsiObjectIdentity>>.Success(new[]
                {
                    CsiObjectIdentity.Create(CsiObjectTypes.Frame, "B2", "Beam 2", "L1")
                }));
            var useCase = new ExportSelectedObjectConnectivityUseCase(tableService, identityResolver);

            OperationResult<PreparedTableExport> result = await useCase.ExecuteAsync(new ObjectConnectivityRequest
            {
                TableName = "Beam Object Connectivity",
                DisplayName = "Beam Object Connectivity",
                ObjectCategory = CsiObjectTypes.Frame
            });

            result.IsSuccess.Should().BeTrue();
            result.Data.RecordCount.Should().Be(1);
            result.Data.Rows[0][0].Should().Be("B2");
        }

        [Fact]
        public async Task ExecuteAsync_surfaces_identity_resolution_failure()
        {
            var tableService = Substitute.For<IEtabsDatabaseTableService>();
            var identityResolver = Substitute.For<ISelectedObjectIdentityResolver>();
            identityResolver.ResolveSelectedObjects().Returns(
                OperationResult<IReadOnlyList<CsiObjectIdentity>>.Failure("selection failed"));
            var useCase = new ExportSelectedObjectConnectivityUseCase(tableService, identityResolver);

            OperationResult<PreparedTableExport> result = await useCase.ExecuteAsync(new ObjectConnectivityRequest
            {
                TableName = "Point Object Connectivity",
                DisplayName = "Point Object Connectivity",
                ObjectCategory = CsiObjectTypes.Point
            });

            result.IsSuccess.Should().BeFalse();
            result.Message.Should().Be("selection failed");
        }
    }
}
