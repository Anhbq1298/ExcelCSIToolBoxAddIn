using System;
using System.Collections.Generic;
using ExcelCSIToolBoxAddIn.Adapters;
using ExcelCSIToolBoxAddIn.Common.Results;
using ExcelCSIToolBoxAddIn.Core.Geometry;
using ExcelCSIToolBoxAddIn.Data;
using ExcelCSIToolBoxAddIn.Data.DTOs;
using ExcelCSIToolBoxAddIn.Data.Models;
using ExcelCSIToolBoxAddIn.Infrastructure.CSISapModel;
using SAP2000v1;

namespace ExcelCSIToolBoxAddIn.Infrastructure.Sap2000
{
    /// <summary>
    /// SAP2000 adapter that exposes the same toolbox contract used by ETABS.
    /// The application/use-case layer stays shared; only the CSI API binding differs here.
    /// </summary>
    public class Sap2000ConnectionService : ICSISapModelConnectionService
    {
        private readonly ICSISapModelConnectionAdapter<SAP2000v1.cSapModel> _connectionAdapter;

        public Sap2000ConnectionService()
            : this(CSISapModelConnectionAdapterFactory.CreateSap2000())
        {
        }

        public Sap2000ConnectionService(ICsiModelAdapter modelAdapter)
            : this(CSISapModelConnectionAdapterFactory.CreateSap2000(modelAdapter))
        {
        }

        private Sap2000ConnectionService(ICSISapModelConnectionAdapter<SAP2000v1.cSapModel> connectionAdapter)
        {
            _connectionAdapter = connectionAdapter ?? throw new ArgumentNullException(nameof(connectionAdapter));
        }

        public string ProductName => "SAP2000";

        public OperationResult<CSISapModelConnectionInfoDTO> TryAttachToRunningInstance()
        {
            return _connectionAdapter.TryAttachToRunningInstance();
        }

        public OperationResult<CSISapModelConnectionInfoDTO> GetCurrentConnection()
        {
            return _connectionAdapter.GetCurrentConnection();
        }

        public OperationResult CloseCurrentInstance()
        {
            return _connectionAdapter.CloseCurrentInstance();
        }

        public OperationResult SelectPointsByUniqueNames(IReadOnlyList<string> uniqueNames)
        {
            var sapModelResult = EnsureSap2000SapModel();
            if (!sapModelResult.IsSuccess)
            {
                return OperationResult.Failure(sapModelResult.Message);
            }

            return CSISapModelPointObjectService.SelectPointsByUniqueNames(
                uniqueNames,
                ProductName,
                sapModelResult.Data,
                sapModel => sapModel.SelectObj.ClearSelection(),
                (sapModel, name) => sapModel.PointObj.SetSelected(name, true, SAP2000v1.eItemType.Objects),
                RefreshView);
        }

        public OperationResult SelectFramesByUniqueNames(IReadOnlyList<string> uniqueNames)
        {
            var sapModelResult = EnsureSap2000SapModel();
            if (!sapModelResult.IsSuccess)
            {
                return OperationResult.Failure(sapModelResult.Message);
            }

            return CSISapModelFrameObjectService.SelectFramesByUniqueNames(
                uniqueNames,
                ProductName,
                sapModelResult.Data,
                sapModel => sapModel.SelectObj.ClearSelection(),
                (sapModel, name) => sapModel.FrameObj.SetSelected(name, true, SAP2000v1.eItemType.Objects),
                RefreshView);
        }

        public OperationResult<CSISapModelAddPointsResultDTO> AddPointsByCartesian(IReadOnlyList<CSISapModelPointCartesianInput> pointInputs)
        {
            var sapModelResult = EnsureSap2000SapModel();
            if (!sapModelResult.IsSuccess)
            {
                return OperationResult<CSISapModelAddPointsResultDTO>.Failure(sapModelResult.Message);
            }

            return CSISapModelPointObjectService.AddPointsByCartesian(
                pointInputs,
                ProductName,
                sapModelResult.Data,
                (SAP2000v1.cSapModel sapModel, CSISapModelPointCartesianInput pointInput, ref string assignedName, string requestedUniqueName) =>
                    sapModel.PointObj.AddCartesian(
                        pointInput.X,
                        pointInput.Y,
                        pointInput.Z,
                        ref assignedName,
                        requestedUniqueName,
                        "Global",
                        true,
                        0),
                RefreshView);
        }

        public OperationResult<CSISapModelAddFramesResultDTO> AddFramesByCoordinates(IReadOnlyList<CSISapModelFrameByCoordInput> frameInputs)
        {
            var sapModelResult = EnsureSap2000SapModel();
            if (!sapModelResult.IsSuccess)
            {
                return OperationResult<CSISapModelAddFramesResultDTO>.Failure(sapModelResult.Message);
            }

            return CSISapModelFrameObjectService.AddFramesByCoordinates(
                frameInputs,
                ProductName,
                sapModelResult.Data,
                (SAP2000v1.cSapModel sapModel, CSISapModelFrameByCoordInput frameInput, ref string createdName, string sectionName, string userName) =>
                    sapModel.FrameObj.AddByCoord(
                        frameInput.Xi,
                        frameInput.Yi,
                        frameInput.Zi,
                        frameInput.Xj,
                        frameInput.Yj,
                        frameInput.Zj,
                        ref createdName,
                        sectionName,
                        userName,
                        "Global"),
                RefreshView);
        }

        public OperationResult<CSISapModelAddFramesResultDTO> AddFramesByPoint(IReadOnlyList<CSISapModelFrameByPointInput> frameInputs)
        {
            var sapModelResult = EnsureSap2000SapModel();
            if (!sapModelResult.IsSuccess)
            {
                return OperationResult<CSISapModelAddFramesResultDTO>.Failure(sapModelResult.Message);
            }

            return CSISapModelFrameObjectService.AddFramesByPoint(
                frameInputs,
                ProductName,
                sapModelResult.Data,
                (SAP2000v1.cSapModel sapModel, CSISapModelFrameByPointInput frameInput, ref string createdName, string sectionName, string userName) =>
                    sapModel.FrameObj.AddByPoint(
                        frameInput.Point1Name,
                        frameInput.Point2Name,
                        ref createdName,
                        sectionName,
                        userName),
                RefreshView);
        }

        public OperationResult<IReadOnlyList<CSISapModelPointDataDTO>> GetSelectedPointsFromActiveModel()
        {
            var sapModelResult = EnsureSap2000SapModel();
            if (!sapModelResult.IsSuccess)
            {
                return OperationResult<IReadOnlyList<CSISapModelPointDataDTO>>.Failure(sapModelResult.Message);
            }

            var pointsResult = CSISapModelPointObjectService.GetSelectedPointsFromActiveModel(
                ProductName,
                sapModelResult.Data,
                (SAP2000v1.cSapModel sapModel, ref int numberItems, ref int[] objectTypes, ref string[] objectNames) =>
                    sapModel.SelectObj.GetSelected(ref numberItems, ref objectTypes, ref objectNames),
                (SAP2000v1.cSapModel sapModel, string pointName, ref double x, ref double y, ref double z) =>
                    sapModel.PointObj.GetCoordCartesian(pointName, ref x, ref y, ref z, "Global"),
                null);
            return pointsResult;
        }

        public OperationResult<IReadOnlyList<string>> GetSelectedFramesFromActiveModel()
        {
            var sapModelResult = EnsureSap2000SapModel();
            if (!sapModelResult.IsSuccess)
            {
                return OperationResult<IReadOnlyList<string>>.Failure(sapModelResult.Message);
            }

            var framesResult = CSISapModelFrameObjectService.GetSelectedFramesFromActiveModel(
                ProductName,
                sapModelResult.Data,
                (SAP2000v1.cSapModel sapModel, ref int numberItems, ref int[] objectTypes, ref string[] objectNames) =>
                    sapModel.SelectObj.GetSelected(ref numberItems, ref objectTypes, ref objectNames));
            return framesResult;
        }

        public OperationResult AddSteelISections(IReadOnlyList<CSISapModelSteelISectionInput> inputs)
        {
            var sapModelResult = EnsureSap2000SapModel();
            if (!sapModelResult.IsSuccess)
            {
                return OperationResult.Failure(sapModelResult.Message);
            }

            var sapModel = sapModelResult.Data;
            var result = CSISapModelSectionPropertiesService.AddSteelISections(
                sapModel,
                inputs);

            return result;
        }

        public OperationResult AddSteelChannelSections(IReadOnlyList<CSISapModelSteelChannelSectionInput> inputs)
        {
            var sapModelResult = EnsureSap2000SapModel();
            if (!sapModelResult.IsSuccess)
            {
                return OperationResult.Failure(sapModelResult.Message);
            }

            var sapModel = sapModelResult.Data;
            var result = CSISapModelSectionPropertiesService.AddSteelChannelSections(
                sapModel,
                inputs);

            return result;
        }

        public OperationResult AddSteelAngleSections(IReadOnlyList<CSISapModelSteelAngleSectionInput> inputs)
        {
            var sapModelResult = EnsureSap2000SapModel();
            if (!sapModelResult.IsSuccess)
            {
                return OperationResult.Failure(sapModelResult.Message);
            }

            var sapModel = sapModelResult.Data;
            var result = CSISapModelSectionPropertiesService.AddSteelAngleSections(
                sapModel,
                inputs);

            return result;
        }

        public OperationResult AddSteelPipeSections(IReadOnlyList<CSISapModelSteelPipeSectionInput> inputs)
        {
            var sapModelResult = EnsureSap2000SapModel();
            if (!sapModelResult.IsSuccess)
            {
                return OperationResult.Failure(sapModelResult.Message);
            }

            var sapModel = sapModelResult.Data;
            var result = CSISapModelSectionPropertiesService.AddSteelPipeSections(
                sapModel,
                inputs);

            return result;
        }

        public OperationResult AddSteelTubeSections(IReadOnlyList<CSISapModelSteelTubeSectionInput> inputs)
        {
            var sapModelResult = EnsureSap2000SapModel();
            if (!sapModelResult.IsSuccess)
            {
                return OperationResult.Failure(sapModelResult.Message);
            }

            var sapModel = sapModelResult.Data;
            var result = CSISapModelSectionPropertiesService.AddSteelTubeSections(
                sapModel,
                inputs);

            return result;
        }

        public OperationResult AddConcreteRectangleSections(IReadOnlyList<CSISapModelConcreteRectangleSectionInput> inputs)
        {
            var sapModelResult = EnsureSap2000SapModel();
            if (!sapModelResult.IsSuccess)
            {
                return OperationResult.Failure(sapModelResult.Message);
            }

            var sapModel = sapModelResult.Data;
            var result = CSISapModelSectionPropertiesService.AddConcreteRectangleSections(
                sapModel,
                inputs);

            return result;
        }

        public OperationResult AddConcreteCircleSections(IReadOnlyList<CSISapModelConcreteCircleSectionInput> inputs)
        {
            var sapModelResult = EnsureSap2000SapModel();
            if (!sapModelResult.IsSuccess)
            {
                return OperationResult.Failure(sapModelResult.Message);
            }

            var sapModel = sapModelResult.Data;
            var result = CSISapModelSectionPropertiesService.AddConcreteCircleSections(
                sapModel,
                inputs);

            return result;
        }

        public OperationResult CreateShellAreasFromSelectedFrames(
            string propertyName,
            ShellCreationTolerances tolerances)
        {
            var sapModelResult = EnsureSap2000SapModel();
            if (!sapModelResult.IsSuccess)
            {
                return OperationResult.Failure(sapModelResult.Message);
            }

            var shellResult = CSISapModelShellObjectService.CreateShellAreasFromSelectedFrames(
                sapModelResult.Data,
                "SAP2000",
                propertyName,
                tolerances,
                sapModel => sapModel.SetPresentUnits(SAP2000v1.eUnits.kN_m_C),
                (SAP2000v1.cSapModel sapModel, ref int numberItems, ref int[] objectTypes, ref string[] objectNames) =>
                    sapModel.SelectObj.GetSelected(ref numberItems, ref objectTypes, ref objectNames),
                (SAP2000v1.cSapModel sapModel, string frameName, ref string point1Name, ref string point2Name) =>
                    sapModel.FrameObj.GetPoints(frameName, ref point1Name, ref point2Name),
                (SAP2000v1.cSapModel sapModel, string pointName, ref double x, ref double y, ref double z) =>
                    sapModel.PointObj.GetCoordCartesian(pointName, ref x, ref y, ref z, "Global"),
                (SAP2000v1.cSapModel sapModel, int nodeCount, ref double[] x, ref double[] y, ref double[] z, ref string areaName, string propName) =>
                    sapModel.AreaObj.AddByCoord(nodeCount, ref x, ref y, ref z, ref areaName, propName, string.Empty, "Global"),
                RefreshView);
            return shellResult;
        }

        private OperationResult<SAP2000v1.cSapModel> EnsureSap2000SapModel()
        {
            return _connectionAdapter.EnsureSapModel();
        }

        public OperationResult<IReadOnlyList<ExcelCSIToolBoxAddIn.Data.DTOs.CSISapModelLoadCombinationDTO>> GetLoadCombinations()
        {
            var sapModelResult = EnsureSap2000SapModel();
            if (!sapModelResult.IsSuccess)
            {
                var errorResult = OperationResult<IReadOnlyList<ExcelCSIToolBoxAddIn.Data.DTOs.CSISapModelLoadCombinationDTO>>.Failure(sapModelResult.Message);
                return errorResult;
            }

            var comboResult = Infrastructure.CSISapModel.LoadCombinationService.CSISapModelLoadCombinationService.GetLoadCombinations(
                sapModelResult.Data,
                (SAP2000v1.cSapModel sapModel, ref int numberNames, ref string[] names) =>
                    sapModel.RespCombo.GetNameList(ref numberNames, ref names),
                (SAP2000v1.cSapModel sapModel, string name) =>
                {
                    int type = 0;
                    sapModel.RespCombo.GetTypeOAPI(name, ref type);
                    switch (type)
                    {
                        case 0: return "Linear Add";
                        case 1: return "Envelope";
                        case 2: return "Absolute Add";
                        case 3: return "SRSS";
                        case 4: return "Range Add";
                        default: return type.ToString();
                    }
                });
            
            return comboResult;
        }

        public OperationResult<IReadOnlyList<ExcelCSIToolBoxAddIn.Data.DTOs.LoadCombinationItemDTO>> GetLoadCombinationDetails(string combinationName)
        {
            var sapModelResult = EnsureSap2000SapModel();
            if (!sapModelResult.IsSuccess)
            {
                var errorResult = OperationResult<IReadOnlyList<ExcelCSIToolBoxAddIn.Data.DTOs.LoadCombinationItemDTO>>.Failure(sapModelResult.Message);
                return errorResult;
            }

            var detailsResult = Infrastructure.CSISapModel.LoadCombinationService.CSISapModelLoadCombinationService.GetLoadCombinationDetails(
                sapModelResult.Data,
                combinationName,
                (SAP2000v1.cSapModel sapModel, string name, ref int numberItems, ref string[] caseNames, ref int[] caseTypes, ref double[] scaleFactors) =>
                {
                    SAP2000v1.eCNameType[] cTypes = null;
                    int ret = sapModel.RespCombo.GetCaseList(name, ref numberItems, ref cTypes, ref caseNames, ref scaleFactors);
                    if (cTypes != null)
                    {
                        caseTypes = new int[cTypes.Length];
                        for (int i = 0; i < cTypes.Length; i++)
                        {
                            caseTypes[i] = (int)cTypes[i];
                        }
                    }
                    return ret;
                },
                (SAP2000v1.cSapModel sapModel, string caseName, int typeCode) =>
                {
                    if (typeCode == 0) // Load Case
                    {
                        SAP2000v1.eLoadCaseType caseType = SAP2000v1.eLoadCaseType.LinearStatic;
                        int subType = 0;
                        int ret = sapModel.LoadCases.GetTypeOAPI(caseName, ref caseType, ref subType);
                        if (ret == 0)
                        {
                            switch (caseType)
                            {
                                case SAP2000v1.eLoadCaseType.LinearStatic: return "Linear Static";
                                case SAP2000v1.eLoadCaseType.NonlinearStatic: return "Nonlinear Static";
                                case SAP2000v1.eLoadCaseType.Modal: return "Modal";
                                case SAP2000v1.eLoadCaseType.ResponseSpectrum: return "Response Spectrum";
                                case SAP2000v1.eLoadCaseType.LinearHistory: return "Linear History";
                                case SAP2000v1.eLoadCaseType.NonlinearHistory: return "Nonlinear History";
                                case SAP2000v1.eLoadCaseType.LinearDynamic: return "Linear Dynamic";
                                case SAP2000v1.eLoadCaseType.NonlinearDynamic: return "Nonlinear Dynamic";
                                case SAP2000v1.eLoadCaseType.MovingLoad: return "Moving Load";
                                case SAP2000v1.eLoadCaseType.Buckling: return "Buckling";
                                case SAP2000v1.eLoadCaseType.SteadyState: return "Steady State";
                                case SAP2000v1.eLoadCaseType.PowerSpectralDensity: return "Power Spectral Density";
                                case SAP2000v1.eLoadCaseType.LinearStaticMultiStep: return "Linear Static Multi-Step";
                                case SAP2000v1.eLoadCaseType.HyperStatic: return "Hyper Static";
                                default: return caseType.ToString();
                            }
                        }
                        return "Load Case";
                    }
                    else
                    {
                        return "Load Combo";
                    }
                });
            
            return detailsResult;
        }

        public OperationResult DeleteLoadCombinations(IReadOnlyList<string> loadCombinationNames)
        {
            var sapModelResult = EnsureSap2000SapModel();
            if (!sapModelResult.IsSuccess)
            {
                return OperationResult.Failure(sapModelResult.Message);
            }

            var result = Infrastructure.CSISapModel.LoadCombinationService.CSISapModelLoadCombinationService.DeleteLoadCombinations(
                sapModelResult.Data,
                loadCombinationNames,
                (SAP2000v1.cSapModel sapModel, string name) => sapModel.RespCombo.Delete(name));
            
            if (result.IsSuccess)
            {
                RefreshView(sapModelResult.Data);
            }

            return result;
        }

        public OperationResult<IReadOnlyList<ExcelCSIToolBoxAddIn.Data.DTOs.CSISapModelLoadPatternDTO>> GetLoadPatterns()
        {
            var sapModelResult = EnsureSap2000SapModel();
            if (!sapModelResult.IsSuccess)
            {
                var errorResult = OperationResult<IReadOnlyList<ExcelCSIToolBoxAddIn.Data.DTOs.CSISapModelLoadPatternDTO>>.Failure(sapModelResult.Message);
                return errorResult;
            }

            var patternResult = Infrastructure.CSISapModel.LoadPatternService.CSISapModelLoadPatternService.GetLoadPatterns(
                sapModelResult.Data,
                (SAP2000v1.cSapModel sapModel, ref int numberNames, ref string[] names) =>
                    sapModel.LoadPatterns.GetNameList(ref numberNames, ref names),
                (SAP2000v1.cSapModel sapModel, string name) =>
                {
                    SAP2000v1.eLoadPatternType type = SAP2000v1.eLoadPatternType.Dead;
                    sapModel.LoadPatterns.GetLoadType(name, ref type);
                    return type.ToString();
                });
            
            return patternResult;
        }

        public OperationResult DeleteLoadPatterns(IReadOnlyList<string> loadPatternNames)
        {
            var sapModelResult = EnsureSap2000SapModel();
            if (!sapModelResult.IsSuccess)
            {
                return OperationResult.Failure(sapModelResult.Message);
            }

            var result = Infrastructure.CSISapModel.LoadPatternService.CSISapModelLoadPatternService.DeleteLoadPatterns(
                sapModelResult.Data,
                loadPatternNames,
                (SAP2000v1.cSapModel sapModel, string name) => sapModel.LoadPatterns.Delete(name));
            
            if (result.IsSuccess)
            {
                RefreshView(sapModelResult.Data);
            }

            return result;
        }

        public OperationResult<IReadOnlyList<CSISapModelFrameSectionDTO>> GetFrameSections()
        {
            var sapModelResult = EnsureSap2000SapModel();
            if (!sapModelResult.IsSuccess)
            {
                return OperationResult<IReadOnlyList<CSISapModelFrameSectionDTO>>.Failure(sapModelResult.Message);
            }

            var sapModel = sapModelResult.Data;
            int numberNames = 0;
            string[] names = null;
            int ret = sapModel.PropFrame.GetNameList(ref numberNames, ref names);

            if (ret != 0 || names == null)
            {
                return OperationResult<IReadOnlyList<CSISapModelFrameSectionDTO>>.Failure("Failed to get frame section names from SAP2000.");
            }

            var list = new List<CSISapModelFrameSectionDTO>();
            for (int i = 0; i < numberNames; i++)
            {
                SAP2000v1.eFramePropType propType = SAP2000v1.eFramePropType.I;
                sapModel.PropFrame.GetTypeOAPI(names[i], ref propType);

                FrameSectionShapeType shapeType = FrameSectionShapeType.Unknown;
                switch (propType)
                {
                    case SAP2000v1.eFramePropType.I: shapeType = FrameSectionShapeType.I; break;
                    case SAP2000v1.eFramePropType.Channel: shapeType = FrameSectionShapeType.Channel; break;
                    case SAP2000v1.eFramePropType.T: shapeType = FrameSectionShapeType.T; break;
                    case SAP2000v1.eFramePropType.Angle: shapeType = FrameSectionShapeType.Angle; break;
                    case SAP2000v1.eFramePropType.Box: shapeType = FrameSectionShapeType.Tube; break;
                    case SAP2000v1.eFramePropType.Pipe: shapeType = FrameSectionShapeType.Pipe; break;
                    case SAP2000v1.eFramePropType.Rectangular: shapeType = FrameSectionShapeType.Rectangular; break;
                    case SAP2000v1.eFramePropType.Circle: shapeType = FrameSectionShapeType.Circular; break;
                    default: shapeType = FrameSectionShapeType.General; break;
                }

                list.Add(new CSISapModelFrameSectionDTO
                {
                    Name = names[i],
                    ShapeType = shapeType
                });
            }

            return OperationResult<IReadOnlyList<CSISapModelFrameSectionDTO>>.Success(list);
        }

        public OperationResult<CSISapModelFrameSectionDetailDTO> GetFrameSectionDetail(string sectionName)
        {
            var sapModelResult = EnsureSap2000SapModel();
            if (!sapModelResult.IsSuccess)
            {
                return OperationResult<CSISapModelFrameSectionDetailDTO>.Failure(sapModelResult.Message);
            }

            var sapModel = sapModelResult.Data;
            SAP2000v1.eFramePropType propType = SAP2000v1.eFramePropType.I;
            int ret = sapModel.PropFrame.GetTypeOAPI(sectionName, ref propType);
            if (ret != 0) return OperationResult<CSISapModelFrameSectionDetailDTO>.Failure("Section not found.");

            var detail = new CSISapModelFrameSectionDetailDTO
            {
                Name = sectionName,
                Dimensions = new Dictionary<string, double>()
            };

            string fileName = "";
            string matProp = "";
            int color = 0;
            string notes = "";
            string guid = "";
            double t3 = 0, t2 = 0, tf = 0, tw = 0, t2b = 0, tfb = 0;

            switch (propType)
            {
                case SAP2000v1.eFramePropType.Pipe:
                    detail.ShapeType = FrameSectionShapeType.Pipe;
                    sapModel.PropFrame.GetPipe(sectionName, ref fileName, ref matProp, ref t3, ref tw, ref color, ref notes, ref guid);
                    detail.Dimensions["Outside diameter ( t3 )"] = t3;
                    detail.Dimensions["Wall thickness ( tw )"] = tw;
                    detail.Color = color;
                    detail.Notes = notes;
                    break;
                case SAP2000v1.eFramePropType.I:
                    detail.ShapeType = FrameSectionShapeType.I;
                    sapModel.PropFrame.GetISection(sectionName, ref fileName, ref matProp, ref t3, ref t2, ref tf, ref tw, ref t2b, ref tfb, ref color, ref notes, ref guid);
                    detail.Dimensions["Total depth ( t3 )"] = t3;
                    detail.Dimensions["Top flange width ( t2 )"] = t2;
                    detail.Dimensions["Top flange thickness ( tf )"] = tf;
                    detail.Dimensions["Web thickness ( tw )"] = tw;
                    detail.Dimensions["Bottom flange width ( t2b )"] = t2b;
                    detail.Dimensions["Bottom flange thickness ( tfb )"] = tfb;
                    detail.Color = color;
                    detail.Notes = notes;
                    break;
                case SAP2000v1.eFramePropType.Channel:
                    detail.ShapeType = FrameSectionShapeType.Channel;
                    sapModel.PropFrame.GetChannel(sectionName, ref fileName, ref matProp, ref t3, ref t2, ref tf, ref tw, ref color, ref notes, ref guid);
                    detail.Dimensions["Total depth ( t3 )"] = t3;
                    detail.Dimensions["Flange width ( t2 )"] = t2;
                    detail.Dimensions["Flange thickness ( tf )"] = tf;
                    detail.Dimensions["Web thickness ( tw )"] = tw;
                    detail.Color = color;
                    detail.Notes = notes;
                    break;
                case SAP2000v1.eFramePropType.Angle:
                    detail.ShapeType = FrameSectionShapeType.Angle;
                    sapModel.PropFrame.GetAngle(sectionName, ref fileName, ref matProp, ref t3, ref t2, ref tf, ref tw, ref color, ref notes, ref guid);
                    detail.Dimensions["Total depth ( t3 )"] = t3;
                    detail.Dimensions["Flange width ( t2 )"] = t2;
                    detail.Dimensions["Flange thickness ( tf )"] = tf;
                    detail.Dimensions["Web thickness ( tw )"] = tw;
                    detail.Color = color;
                    detail.Notes = notes;
                    break;
                case SAP2000v1.eFramePropType.Rectangular:
                    detail.ShapeType = FrameSectionShapeType.Rectangular;
                    sapModel.PropFrame.GetRectangle(sectionName, ref fileName, ref matProp, ref t3, ref t2, ref color, ref notes, ref guid);
                    detail.Dimensions["Depth ( t3 )"] = t3;
                    detail.Dimensions["Width ( t2 )"] = t2;
                    detail.Color = color;
                    detail.Notes = notes;
                    break;
                case SAP2000v1.eFramePropType.Circle:
                    detail.ShapeType = FrameSectionShapeType.Circular;
                    sapModel.PropFrame.GetCircle(sectionName, ref fileName, ref matProp, ref t3, ref color, ref notes, ref guid);
                    detail.Dimensions["Diameter ( t3 )"] = t3;
                    detail.Color = color;
                    detail.Notes = notes;
                    break;
                case SAP2000v1.eFramePropType.Box:
                    detail.ShapeType = FrameSectionShapeType.Tube;
                    sapModel.PropFrame.GetTube(sectionName, ref fileName, ref matProp, ref t3, ref t2, ref tf, ref tw, ref color, ref notes, ref guid);
                    detail.Dimensions["Total depth ( t3 )"] = t3;
                    detail.Dimensions["Flange width ( t2 )"] = t2;
                    detail.Dimensions["Flange thickness ( tf )"] = tf;
                    detail.Dimensions["Web thickness ( tw )"] = tw;
                    detail.Color = color;
                    detail.Notes = notes;
                    break;
                default:
                    detail.ShapeType = FrameSectionShapeType.Unknown;
                    break;
            }

            detail.MaterialName = matProp;
            return OperationResult<CSISapModelFrameSectionDetailDTO>.Success(detail);
        }

        private static OperationResult RefreshView(SAP2000v1.cSapModel sapModel)
        {
            int refreshResult = sapModel.View.RefreshView(0, false);
            if (refreshResult != 0)
            {
                return OperationResult.Failure($"SAP2000 model changed successfully, but View.RefreshView failed (return code {refreshResult}).");
            }

            return OperationResult.Success();
        }

    }
}


