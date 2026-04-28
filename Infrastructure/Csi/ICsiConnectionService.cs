using System.Collections.Generic;
using ExcelCSIToolBoxAddIn.Common.Results;
using ExcelCSIToolBoxAddIn.Core.Geometry;

namespace ExcelCSIToolBoxAddIn.Infrastructure.Csi
{
    public interface ICsiConnectionService
    {
        string ProductName { get; }

        OperationResult<CsiConnectionInfo> TryAttachToRunningInstance();

        OperationResult<CsiConnectionInfo> GetCurrentConnection();

        OperationResult CloseCurrentInstance();

        OperationResult SelectPointsByUniqueNames(IReadOnlyList<string> uniqueNames);
        OperationResult SelectFramesByUniqueNames(IReadOnlyList<string> uniqueNames);

        // pointInputs must be executed exactly in the given order.
        // Duplicate rows are valid and must not be merged or de-duplicated.
        OperationResult<CsiAddPointsResult> AddPointsByCartesian(IReadOnlyList<CsiPointCartesianInput> pointInputs);
        OperationResult<CsiAddFramesResult> AddFramesByCoordinates(IReadOnlyList<CsiFrameByCoordInput> frameInputs);
        OperationResult<CsiAddFramesResult> AddFramesByPoint(IReadOnlyList<CsiFrameByPointInput> frameInputs);

        OperationResult<IReadOnlyList<CsiPointData>> GetSelectedPointsFromActiveModel();
        OperationResult<IReadOnlyList<string>> GetSelectedFramesFromActiveModel();

        OperationResult AddSteelISections(IReadOnlyList<CsiSteelISectionInput> inputs);
        OperationResult AddSteelChannelSections(IReadOnlyList<CsiSteelChannelSectionInput> inputs);
        OperationResult AddSteelAngleSections(IReadOnlyList<CsiSteelAngleSectionInput> inputs);
        OperationResult AddSteelPipeSections(IReadOnlyList<CsiSteelPipeSectionInput> inputs);
        OperationResult AddSteelTubeSections(IReadOnlyList<CsiSteelTubeSectionInput> inputs);

        OperationResult AddConcreteRectangleSections(IReadOnlyList<CsiConcreteRectangleSectionInput> inputs);
        OperationResult AddConcreteCircleSections(IReadOnlyList<CsiConcreteCircleSectionInput> inputs);

        OperationResult CreateShellAreasFromSelectedFrames(
            string propertyName,
            ShellCreationTolerances tolerances);
    }
}
