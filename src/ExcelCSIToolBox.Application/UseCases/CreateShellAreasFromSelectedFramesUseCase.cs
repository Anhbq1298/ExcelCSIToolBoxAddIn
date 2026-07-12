using ExcelCSIToolBox.Core.Common.Results;
using ExcelCSIToolBox.Core.Geometry;
using ExcelCSIToolBox.Core.Abstractions.CSI;

namespace ExcelCSIToolBox.Application.UseCases
{
    public class CreateShellAreasFromSelectedFramesUseCase
    {
        private readonly ICSISapModelConnectionService _connectionService;

        public CreateShellAreasFromSelectedFramesUseCase(ICSISapModelConnectionService connectionService)
        {
            _connectionService = connectionService;
        }

        public OperationResult Execute(string propertyName)
        {
            var propName = string.IsNullOrWhiteSpace(propertyName)
                ? "Default"
                : propertyName.Trim();

            return _connectionService.CreateShellAreasFromSelectedFrames(
                propName,
                new ShellCreationTolerances());
        }
    }
}

