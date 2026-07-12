using System.Collections.Generic;
using ExcelCSIToolBox.Application.UseCases;
using ExcelCSIToolBox.Core.Abstractions.CSI;
using ExcelCSIToolBox.Core.Common.Results;
using ExcelCSIToolBox.Data.DTOs.CSI;
using FluentAssertions;
using NSubstitute;
using Xunit;

namespace ExcelCSIToolBox.Tests
{
    public class GetFrameSectionsUseCaseTests
    {
        [Fact]
        public void Execute_returns_sections_from_connection_service()
        {
            var service = Substitute.For<ICSISapModelConnectionService>();
            var sections = new List<CSISapModelFrameSectionDTO>
            {
                new CSISapModelFrameSectionDTO { Name = "W12X26", MaterialName = "A992" }
            };
            service.GetFrameSections().Returns(OperationResult<IReadOnlyList<CSISapModelFrameSectionDTO>>.Success(sections));
            var useCase = new GetFrameSectionsUseCase(service);

            OperationResult<IReadOnlyList<CSISapModelFrameSectionDTO>> result = useCase.Execute();

            result.IsSuccess.Should().BeTrue();
            result.Data.Should().ContainSingle(section => section.Name == "W12X26");
        }

        [Fact]
        public void Execute_surfaces_failure_from_connection_service()
        {
            var service = Substitute.For<ICSISapModelConnectionService>();
            service.GetFrameSections().Returns(OperationResult<IReadOnlyList<CSISapModelFrameSectionDTO>>.Failure("not attached"));
            var useCase = new GetFrameSectionsUseCase(service);

            OperationResult<IReadOnlyList<CSISapModelFrameSectionDTO>> result = useCase.Execute();

            result.IsSuccess.Should().BeFalse();
            result.Message.Should().Be("not attached");
        }
    }
}
