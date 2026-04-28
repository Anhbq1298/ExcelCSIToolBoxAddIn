using System.Collections.Generic;
using ExcelCSIToolBoxAddIn.Common.Results;
using ExcelCSIToolBoxAddIn.Core.Geometry;

namespace ExcelCSIToolBoxAddIn.Infrastructure.CSISapModel
{
    public interface ICSISapModelConnectionService
    {
        string ProductName { get; }

        OperationResult<CSISapModelConnectionInfo> TryAttachToRunningInstance();

        OperationResult<CSISapModelConnectionInfo> GetCurrentConnection();

        OperationResult CloseCurrentInstance();

        OperationResult SelectPointsByUniqueNames(IReadOnlyList<string> uniqueNames);
        OperationResult SelectFramesByUniqueNames(IReadOnlyList<string> uniqueNames);

        // pointInputs must be executed exactly in the given order.
        // Duplicate rows are valid and must not be merged or de-duplicated.
        OperationResult<CSISapModelAddPointsResult> AddPointsByCartesian(IReadOnlyList<CSISapModelPointCartesianInput> pointInputs);
        OperationResult<CSISapModelAddFramesResult> AddFramesByCoordinates(IReadOnlyList<CSISapModelFrameByCoordInput> frameInputs);
        OperationResult<CSISapModelAddFramesResult> AddFramesByPoint(IReadOnlyList<CSISapModelFrameByPointInput> frameInputs);

        OperationResult<IReadOnlyList<CSISapModelPointData>> GetSelectedPointsFromActiveModel();
        OperationResult<IReadOnlyList<string>> GetSelectedFramesFromActiveModel();

        OperationResult AddSteelISections(IReadOnlyList<CSISapModelSteelISectionInput> inputs);
        OperationResult AddSteelChannelSections(IReadOnlyList<CSISapModelSteelChannelSectionInput> inputs);
        OperationResult AddSteelAngleSections(IReadOnlyList<CSISapModelSteelAngleSectionInput> inputs);
        OperationResult AddSteelPipeSections(IReadOnlyList<CSISapModelSteelPipeSectionInput> inputs);
        OperationResult AddSteelTubeSections(IReadOnlyList<CSISapModelSteelTubeSectionInput> inputs);

        OperationResult AddConcreteRectangleSections(IReadOnlyList<CSISapModelConcreteRectangleSectionInput> inputs);
        OperationResult AddConcreteCircleSections(IReadOnlyList<CSISapModelConcreteCircleSectionInput> inputs);

        OperationResult CreateShellAreasFromSelectedFrames(
            string propertyName,
            ShellCreationTolerances tolerances);
    }
}
