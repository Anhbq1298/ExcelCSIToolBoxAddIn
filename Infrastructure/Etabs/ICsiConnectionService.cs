using System.Collections.Generic;
using ExcelCSIToolBoxAddIn.Common.Results;
using ExcelCSIToolBoxAddIn.Core.Geometry;

namespace ExcelCSIToolBoxAddIn.Infrastructure.Etabs
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
        OperationResult<CsiAddPointsResult> AddPointsByCartesian(IReadOnlyList<EtabsPointCartesianInput> pointInputs);
        OperationResult<CsiAddFramesResult> AddFramesByCoordinates(IReadOnlyList<EtabsFrameByCoordInput> frameInputs);
        OperationResult<CsiAddFramesResult> AddFramesByPoint(IReadOnlyList<EtabsFrameByPointInput> frameInputs);

        OperationResult<IReadOnlyList<CsiPointData>> GetSelectedPointsFromActiveModel();
        OperationResult<IReadOnlyList<string>> GetSelectedFramesFromActiveModel();

        OperationResult AddSteelISections(IReadOnlyList<EtabsSteelISectionInput> inputs);
        OperationResult AddSteelChannelSections(IReadOnlyList<EtabsSteelChannelSectionInput> inputs);
        OperationResult AddSteelAngleSections(IReadOnlyList<EtabsSteelAngleSectionInput> inputs);
        OperationResult AddSteelPipeSections(IReadOnlyList<EtabsSteelPipeSectionInput> inputs);
        OperationResult AddSteelTubeSections(IReadOnlyList<EtabsSteelTubeSectionInput> inputs);

        OperationResult AddConcreteRectangleSections(IReadOnlyList<EtabsConcreteRectangleSectionInput> inputs);
        OperationResult AddConcreteCircleSections(IReadOnlyList<EtabsConcreteCircleSectionInput> inputs);

        OperationResult CreateShellAreasFromSelectedFrames(
            string propertyName,
            ShellCreationTolerances tolerances);
    }
}
