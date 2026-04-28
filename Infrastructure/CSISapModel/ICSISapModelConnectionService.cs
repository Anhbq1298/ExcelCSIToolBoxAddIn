using System.Collections.Generic;
using ExcelCSIToolBoxAddIn.Common.Results;
using ExcelCSIToolBoxAddIn.Core.Geometry;
using ExcelCSIToolBoxAddIn.Data;
using ExcelCSIToolBoxAddIn.Data.DTOs;
using ExcelCSIToolBoxAddIn.Data.Models;


namespace ExcelCSIToolBoxAddIn.Infrastructure.CSISapModel
{
    public interface ICSISapModelConnectionService
    {
        string ProductName { get; }

        OperationResult<CSISapModelConnectionInfoDTO> TryAttachToRunningInstance();

        OperationResult<CSISapModelConnectionInfoDTO> GetCurrentConnection();

        OperationResult CloseCurrentInstance();

        OperationResult SelectPointsByUniqueNames(IReadOnlyList<string> uniqueNames);
        OperationResult SelectFramesByUniqueNames(IReadOnlyList<string> uniqueNames);

        // pointInputs must be executed exactly in the given order.
        // Duplicate rows are valid and must not be merged or de-duplicated.
        OperationResult<CSISapModelAddPointsResultDTO> AddPointsByCartesian(IReadOnlyList<CSISapModelPointCartesianInput> pointInputs);
        OperationResult<CSISapModelAddFramesResultDTO> AddFramesByCoordinates(IReadOnlyList<CSISapModelFrameByCoordInput> frameInputs);
        OperationResult<CSISapModelAddFramesResultDTO> AddFramesByPoint(IReadOnlyList<CSISapModelFrameByPointInput> frameInputs);

        OperationResult<IReadOnlyList<CSISapModelPointDataDTO>> GetSelectedPointsFromActiveModel();
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

        OperationResult<IReadOnlyList<string>> GetLoadCombinations();
        OperationResult DeleteLoadCombinations(IReadOnlyList<string> loadCombinationNames);

        OperationResult<IReadOnlyList<CSISapModelLoadPatternDTO>> GetLoadPatterns();
        OperationResult DeleteLoadPatterns(IReadOnlyList<string> loadPatternNames);
    }
}
