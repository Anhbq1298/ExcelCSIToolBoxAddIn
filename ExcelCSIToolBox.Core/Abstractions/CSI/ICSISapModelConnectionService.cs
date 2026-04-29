using System.Collections.Generic;
using ExcelCSIToolBox.Core.Common.Results;
using ExcelCSIToolBox.Core.Geometry;
using ExcelCSIToolBox.Core.Models.CSI;
using ExcelCSIToolBox.Data;
using ExcelCSIToolBox.Data.CSISapModel.FrameObject;
using ExcelCSIToolBox.Data.CSISapModel.PointObject;
using ExcelCSIToolBox.Data.DTOs.CSI;
using ExcelCSIToolBox.Data.Models;


namespace ExcelCSIToolBox.Core.Abstractions.CSI
{
    public interface ICSISapModelConnectionService
    {
        string ProductName { get; }

        OperationResult<CSISapModelConnectionInfoDTO> TryAttachToRunningInstance();

        OperationResult<CSISapModelConnectionInfoDTO> GetCurrentConnection();

        OperationResult CloseCurrentInstance();

        OperationResult SelectPointsByUniqueNames(IReadOnlyList<string> uniqueNames);
        OperationResult SelectFramesByUniqueNames(IReadOnlyList<string> uniqueNames);
        OperationResult ClearSelection();

        // pointInputs must be executed exactly in the given order.
        // Duplicate rows are valid and must not be merged or de-duplicated.
        OperationResult<CSISapModelAddPointsResultDTO> AddPointsByCartesian(IReadOnlyList<CSISapModelPointCartesianInput> pointInputs);
        OperationResult<CSISapModelAddFramesResultDTO> AddFramesByCoordinates(IReadOnlyList<CSISapModelFrameByCoordInput> frameInputs);
        OperationResult<CSISapModelAddFramesResultDTO> AddFramesByPoint(IReadOnlyList<CSISapModelFrameByPointInput> frameInputs);
        OperationResult AssignFrameSection(IReadOnlyList<string> frameNames, string sectionName);
        OperationResult AssignFrameDistributedLoad(IReadOnlyList<string> frameNames, string loadPattern, int direction, double value1, double value2);
        OperationResult AssignFramePointLoad(IReadOnlyList<string> frameNames, string loadPattern, int direction, double distance, double value);
        OperationResult DeleteFrameObjects(IReadOnlyList<string> frameNames);
        OperationResult RunAnalysis();
        OperationResult SaveModel();

        OperationResult<IReadOnlyList<CSISapModelPointDataDTO>> GetSelectedPointsFromActiveModel();
        OperationResult<IReadOnlyList<string>> GetPointNames();
        OperationResult<PointObjectInfo> GetPointByName(string pointName);
        OperationResult<PointObjectInfo> GetPointCoordinates(string pointName);
        OperationResult<PointRestraintInfo> GetPointRestraint(string pointName);
        OperationResult<IReadOnlyList<PointLoadInfo>> GetPointLoadForces(string pointName);
        OperationResult SetPointRestraint(IReadOnlyList<string> pointNames, IReadOnlyList<bool> restraints);
        OperationResult SetPointLoadForce(IReadOnlyList<string> pointNames, string loadPattern, IReadOnlyList<double> forceValues, bool replace, string coordinateSystem);
        OperationResult<IReadOnlyList<string>> GetSelectedFramesFromActiveModel();
        OperationResult<IReadOnlyList<string>> GetFrameNames();
        OperationResult<FrameObjectInfo> GetFrameByName(string frameName);
        OperationResult<FrameEndPointInfo> GetFramePoints(string frameName);
        OperationResult<FrameSectionInfo> GetFrameSection(string frameName);
        OperationResult<IReadOnlyList<FrameLoadInfo>> GetFrameDistributedLoads(string frameName);
        OperationResult<IReadOnlyList<FrameLoadInfo>> GetFramePointLoads(string frameName);

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
        OperationResult<IReadOnlyList<string>> GetShellNames();
        OperationResult<CSISapModelShellObjectDTO> GetShellByName(string areaName);
        OperationResult<IReadOnlyList<string>> GetShellPoints(string areaName);
        OperationResult<string> GetShellProperty(string areaName);
        OperationResult<IReadOnlyList<string>> GetSelectedShells();
        OperationResult<IReadOnlyList<CSISapModelShellLoadDTO>> GetShellUniformLoads(string areaName);
        CsiWritePreview PreviewAddShellByPoint(IReadOnlyList<string> pointNames, string propertyName, string userName);
        OperationResult<string> AddShellByPoint(IReadOnlyList<string> pointNames, string propertyName, string userName, bool confirmed);
        CsiWritePreview PreviewAddShellByCoord(IReadOnlyList<CSISapModelShellCoordinateInput> points, string propertyName, string userName, string coordinateSystem);
        OperationResult<string> AddShellByCoord(IReadOnlyList<CSISapModelShellCoordinateInput> points, string propertyName, string userName, string coordinateSystem, bool confirmed);
        CsiWritePreview PreviewAssignShellUniformLoad(IReadOnlyList<string> areaNames, string loadPattern, double value, int direction, bool replace, string coordinateSystem);
        OperationResult AssignShellUniformLoad(IReadOnlyList<string> areaNames, string loadPattern, double value, int direction, bool replace, string coordinateSystem, bool confirmed);
        CsiWritePreview PreviewDeleteShells(IReadOnlyList<string> areaNames);
        OperationResult DeleteShells(IReadOnlyList<string> areaNames, bool confirmed);

        OperationResult<IReadOnlyList<CSISapModelFrameSectionDTO>> GetFrameSections();
        OperationResult<CSISapModelFrameSectionDetailDTO> GetFrameSectionDetail(string sectionName);
        OperationResult UpdateFrameSection(CSISapModelFrameSectionUpdateDTO input);
        OperationResult RenameFrameSection(CSISapModelFrameSectionRenameDTO input);

        OperationResult<IReadOnlyList<ExcelCSIToolBox.Data.DTOs.CSI.CSISapModelLoadCombinationDTO>> GetLoadCombinations();
        OperationResult<IReadOnlyList<ExcelCSIToolBox.Data.DTOs.CSI.LoadCombinationItemDTO>> GetLoadCombinationDetails(string combinationName);
        OperationResult DeleteLoadCombinations(IReadOnlyList<string> loadCombinationNames);

        OperationResult<IReadOnlyList<CSISapModelLoadPatternDTO>> GetLoadPatterns();
        OperationResult DeleteLoadPatterns(IReadOnlyList<string> loadPatternNames);
        OperationResult<CSISapModelStatisticsDTO> GetModelStatistics();
        OperationResult RefreshView(bool zoomAll = false);
        OperationResult SetPresentUnits(int unitsCode);
    }
}


