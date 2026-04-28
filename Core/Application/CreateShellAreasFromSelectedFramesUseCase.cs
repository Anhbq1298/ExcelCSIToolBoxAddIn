using ExcelCSIToolBoxAddIn.Common.Results;
using ExcelCSIToolBoxAddIn.Core.Geometry;
using ExcelCSIToolBoxAddIn.Infrastructure.Csi;

namespace ExcelCSIToolBoxAddIn.Core.Application
{
    public class CreateShellAreasFromSelectedFramesUseCase
    {
        private readonly ICsiConnectionService _connectionService;

        public CreateShellAreasFromSelectedFramesUseCase(ICsiConnectionService connectionService)
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
