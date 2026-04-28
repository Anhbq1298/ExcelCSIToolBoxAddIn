using System;
using System.Collections.Generic;
using ExcelCSIToolBoxAddIn.Adapters;
using ExcelCSIToolBoxAddIn.Common.Results;
using ExcelCSIToolBoxAddIn.Core.Geometry;
using ExcelCSIToolBoxAddIn.Data;
using ExcelCSIToolBoxAddIn.Data.DTOs;
using ExcelCSIToolBoxAddIn.Data.Models;
using ExcelCSIToolBoxAddIn.Infrastructure.CSISapModel;

namespace ExcelCSIToolBoxAddIn.Infrastructure.Etabs
{
    /// <summary>
    /// Infrastructure service that safely attaches to a running ETABS instance.
    /// Stores the latest attached instance so ETABS commands can reuse the same SapModel.
    /// </summary>
    public class EtabsConnectionService : ICSISapModelConnectionService
    {
        private readonly ICSISapModelConnectionAdapter<ETABSv1.cSapModel> _connectionAdapter;

        public EtabsConnectionService()
            : this(CSISapModelConnectionAdapterFactory.CreateEtabs())
        {
        }

        public EtabsConnectionService(ICsiModelAdapter modelAdapter)
            : this(CSISapModelConnectionAdapterFactory.CreateEtabs(modelAdapter))
        {
        }

        private EtabsConnectionService(ICSISapModelConnectionAdapter<ETABSv1.cSapModel> connectionAdapter)
        {
            _connectionAdapter = connectionAdapter ?? throw new ArgumentNullException(nameof(connectionAdapter));
        }

        public string ProductName => "ETABS";

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
            var sapModelResult = EnsureEtabsSapModel();
            if (!sapModelResult.IsSuccess)
            {
                return OperationResult.Failure(sapModelResult.Message);
            }

            var selectResult = CSISapModelPointObjectService.SelectPointsByUniqueNames(
                uniqueNames,
                ProductName,
                sapModelResult.Data,
                sapModel => sapModel.SelectObj.ClearSelection(),
                (sapModel, name) => sapModel.PointObj.SetSelected(name, true, ETABSv1.eItemType.Objects),
                RefreshView);
            return selectResult;
        }

        public OperationResult SelectFramesByUniqueNames(IReadOnlyList<string> uniqueNames)
        {
            var sapModelResult = EnsureEtabsSapModel();
            if (!sapModelResult.IsSuccess)
            {
                return OperationResult.Failure(sapModelResult.Message);
            }

            var selectResult = CSISapModelFrameObjectService.SelectFramesByUniqueNames(
                uniqueNames,
                ProductName,
                sapModelResult.Data,
                sapModel => sapModel.SelectObj.ClearSelection(),
                (sapModel, name) => sapModel.FrameObj.SetSelected(name, true, ETABSv1.eItemType.Objects),
                RefreshView);
            return selectResult;
        }

        public OperationResult<CSISapModelAddPointsResultDTO> AddPointsByCartesian(IReadOnlyList<CSISapModelPointCartesianInput> pointInputs)
        {
            var sapModelResult = EnsureEtabsSapModel();
            if (!sapModelResult.IsSuccess)
            {
                return OperationResult<CSISapModelAddPointsResultDTO>.Failure(sapModelResult.Message);
            }

            var addResult = CSISapModelPointObjectService.AddPointsByCartesian(
                pointInputs,
                ProductName,
                sapModelResult.Data,
                (ETABSv1.cSapModel sapModel, CSISapModelPointCartesianInput pointInput, ref string assignedName, string requestedUniqueName) =>
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
            return addResult;
        }

        public OperationResult<CSISapModelAddFramesResultDTO> AddFramesByCoordinates(IReadOnlyList<CSISapModelFrameByCoordInput> frameInputs)
        {
            var sapModelResult = EnsureEtabsSapModel();
            if (!sapModelResult.IsSuccess)
            {
                return OperationResult<CSISapModelAddFramesResultDTO>.Failure(sapModelResult.Message);
            }

            var addResult = CSISapModelFrameObjectService.AddFramesByCoordinates(
                frameInputs,
                ProductName,
                sapModelResult.Data,
                (ETABSv1.cSapModel sapModel, CSISapModelFrameByCoordInput frameInput, ref string createdName, string sectionName, string userName) =>
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
            return addResult;
        }

        public OperationResult<CSISapModelAddFramesResultDTO> AddFramesByPoint(IReadOnlyList<CSISapModelFrameByPointInput> frameInputs)
        {
            var sapModelResult = EnsureEtabsSapModel();
            if (!sapModelResult.IsSuccess)
            {
                return OperationResult<CSISapModelAddFramesResultDTO>.Failure(sapModelResult.Message);
            }

            var addResult = CSISapModelFrameObjectService.AddFramesByPoint(
                frameInputs,
                ProductName,
                sapModelResult.Data,
                (ETABSv1.cSapModel sapModel, CSISapModelFrameByPointInput frameInput, ref string createdName, string sectionName, string userName) =>
                    sapModel.FrameObj.AddByPoint(
                        frameInput.Point1Name,
                        frameInput.Point2Name,
                        ref createdName,
                        sectionName,
                        userName),
                RefreshView);
            return addResult;
        }

        public OperationResult<IReadOnlyList<CSISapModelPointDataDTO>> GetSelectedPointsFromActiveModel()
        {
            var sapModelResult = EnsureEtabsSapModel();
            if (!sapModelResult.IsSuccess)
            {
                return OperationResult<IReadOnlyList<CSISapModelPointDataDTO>>.Failure(sapModelResult.Message);
            }

            var pointsResult = CSISapModelPointObjectService.GetSelectedPointsFromActiveModel(
                ProductName,
                sapModelResult.Data,
                (ETABSv1.cSapModel sapModel, ref int numberItems, ref int[] objectTypes, ref string[] objectNames) =>
                    sapModel.SelectObj.GetSelected(ref numberItems, ref objectTypes, ref objectNames),
                (ETABSv1.cSapModel sapModel, string pointName, ref double x, ref double y, ref double z) =>
                    sapModel.PointObj.GetCoordCartesian(pointName, ref x, ref y, ref z, "Global"),
                (ETABSv1.cSapModel sapModel, string pointName, ref string pointLabel, ref string pointStory) =>
                    sapModel.PointObj.GetLabelFromName(pointName, ref pointLabel, ref pointStory));
            return pointsResult;
        }

        public OperationResult<IReadOnlyList<string>> GetSelectedFramesFromActiveModel()
        {
            var sapModelResult = EnsureEtabsSapModel();
            if (!sapModelResult.IsSuccess)
            {
                return OperationResult<IReadOnlyList<string>>.Failure(sapModelResult.Message);
            }

            var framesResult = CSISapModelFrameObjectService.GetSelectedFramesFromActiveModel(
                ProductName,
                sapModelResult.Data,
                (ETABSv1.cSapModel sapModel, ref int numberItems, ref int[] objectTypes, ref string[] objectNames) =>
                    sapModel.SelectObj.GetSelected(ref numberItems, ref objectTypes, ref objectNames));
            return framesResult;
        }

        private OperationResult<ETABSv1.cSapModel> EnsureEtabsSapModel()
        {
            return _connectionAdapter.EnsureSapModel();
        }

        public OperationResult AddSteelISections(IReadOnlyList<CSISapModelSteelISectionInput> inputs)
        {
            var sapModelResult = EnsureEtabsSapModel();
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
            var sapModelResult = EnsureEtabsSapModel();
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
            var sapModelResult = EnsureEtabsSapModel();
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
            var sapModelResult = EnsureEtabsSapModel();
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
            var sapModelResult = EnsureEtabsSapModel();
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
            var sapModelResult = EnsureEtabsSapModel();
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
            var sapModelResult = EnsureEtabsSapModel();
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
            var sapModelResult = EnsureEtabsSapModel();
            if (!sapModelResult.IsSuccess)
            {
                return OperationResult.Failure(sapModelResult.Message);
            }

            var shellResult = CSISapModelShellObjectService.CreateShellAreasFromSelectedFrames(
                sapModelResult.Data,
                "ETABS",
                propertyName,
                tolerances,
                sapModel => sapModel.SetPresentUnits(ETABSv1.eUnits.kN_m_C),
                (ETABSv1.cSapModel sapModel, ref int numberItems, ref int[] objectTypes, ref string[] objectNames) =>
                    sapModel.SelectObj.GetSelected(ref numberItems, ref objectTypes, ref objectNames),
                (ETABSv1.cSapModel sapModel, string frameName, ref string point1Name, ref string point2Name) =>
                    sapModel.FrameObj.GetPoints(frameName, ref point1Name, ref point2Name),
                (ETABSv1.cSapModel sapModel, string pointName, ref double x, ref double y, ref double z) =>
                    sapModel.PointObj.GetCoordCartesian(pointName, ref x, ref y, ref z, "Global"),
                (ETABSv1.cSapModel sapModel, int nodeCount, ref double[] x, ref double[] y, ref double[] z, ref string areaName, string propName) =>
                    sapModel.AreaObj.AddByCoord(nodeCount, ref x, ref y, ref z, ref areaName, propName, string.Empty, "Global"),
                RefreshView);
            return shellResult;
        }

        public OperationResult<IReadOnlyList<ExcelCSIToolBoxAddIn.Data.DTOs.CSISapModelLoadCombinationDTO>> GetLoadCombinations()
        {
            var sapModelResult = EnsureEtabsSapModel();
            if (!sapModelResult.IsSuccess)
            {
                var errorResult = OperationResult<IReadOnlyList<ExcelCSIToolBoxAddIn.Data.DTOs.CSISapModelLoadCombinationDTO>>.Failure(sapModelResult.Message);
                return errorResult;
            }

            var comboResult = Infrastructure.CSISapModel.LoadCombinationService.CSISapModelLoadCombinationService.GetLoadCombinations(
                sapModelResult.Data,
                (ETABSv1.cSapModel sapModel, ref int numberNames, ref string[] names) =>
                    sapModel.RespCombo.GetNameList(ref numberNames, ref names),
                (ETABSv1.cSapModel sapModel, string name) =>
                {
                    int type = 0;
                    sapModel.RespCombo.GetTypeOAPI(name, ref type);
                    // Usually 0=Linear Add, 1=Envelope, 2=Absolute Add, 3=SRSS, 4=Range Add
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
            var sapModelResult = EnsureEtabsSapModel();
            if (!sapModelResult.IsSuccess)
            {
                var errorResult = OperationResult<IReadOnlyList<ExcelCSIToolBoxAddIn.Data.DTOs.LoadCombinationItemDTO>>.Failure(sapModelResult.Message);
                return errorResult;
            }

            var detailsResult = Infrastructure.CSISapModel.LoadCombinationService.CSISapModelLoadCombinationService.GetLoadCombinationDetails(
                sapModelResult.Data,
                combinationName,
                (ETABSv1.cSapModel sapModel, string name, ref int numberItems, ref string[] caseNames, ref int[] caseTypes, ref double[] scaleFactors) =>
                {
                    ETABSv1.eCNameType[] cTypes = null;
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
                (ETABSv1.cSapModel sapModel, string caseName, int typeCode) =>
                {
                    if (typeCode == 0) // Load Case
                    {
                        ETABSv1.eLoadCaseType caseType = ETABSv1.eLoadCaseType.LinearStatic;
                        int subType = 0;
                        int ret = sapModel.LoadCases.GetTypeOAPI(caseName, ref caseType, ref subType);
                        if (ret == 0)
                        {
                            switch (caseType)
                            {
                                case ETABSv1.eLoadCaseType.LinearStatic: return "Linear Static";
                                case ETABSv1.eLoadCaseType.NonlinearStatic: return "Nonlinear Static";
                                case ETABSv1.eLoadCaseType.Modal: return "Modal";
                                case ETABSv1.eLoadCaseType.ResponseSpectrum: return "Response Spectrum";
                                case ETABSv1.eLoadCaseType.LinearHistory: return "Linear History";
                                case ETABSv1.eLoadCaseType.NonlinearHistory: return "Nonlinear History";
                                case ETABSv1.eLoadCaseType.LinearDynamic: return "Linear Dynamic";
                                case ETABSv1.eLoadCaseType.NonlinearDynamic: return "Nonlinear Dynamic";
                                case ETABSv1.eLoadCaseType.MovingLoad: return "Moving Load";
                                case ETABSv1.eLoadCaseType.Buckling: return "Buckling";
                                case ETABSv1.eLoadCaseType.SteadyState: return "Steady State";
                                case ETABSv1.eLoadCaseType.PowerSpectralDensity: return "Power Spectral Density";
                                case ETABSv1.eLoadCaseType.LinearStaticMultiStep: return "Linear Static Multi-Step";
                                case ETABSv1.eLoadCaseType.HyperStatic: return "Hyper Static";
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
            var sapModelResult = EnsureEtabsSapModel();
            if (!sapModelResult.IsSuccess)
            {
                return OperationResult.Failure(sapModelResult.Message);
            }

            var result = Infrastructure.CSISapModel.LoadCombinationService.CSISapModelLoadCombinationService.DeleteLoadCombinations(
                sapModelResult.Data,
                loadCombinationNames,
                (ETABSv1.cSapModel sapModel, string name) => sapModel.RespCombo.Delete(name));
            
            if (result.IsSuccess)
            {
                RefreshView(sapModelResult.Data);
            }

            return result;
        }

        public OperationResult<IReadOnlyList<ExcelCSIToolBoxAddIn.Data.DTOs.CSISapModelLoadPatternDTO>> GetLoadPatterns()
        {
            var sapModelResult = EnsureEtabsSapModel();
            if (!sapModelResult.IsSuccess)
            {
                var errorResult = OperationResult<IReadOnlyList<ExcelCSIToolBoxAddIn.Data.DTOs.CSISapModelLoadPatternDTO>>.Failure(sapModelResult.Message);
                return errorResult;
            }

            var patternResult = Infrastructure.CSISapModel.LoadPatternService.CSISapModelLoadPatternService.GetLoadPatterns(
                sapModelResult.Data,
                (ETABSv1.cSapModel sapModel, ref int numberNames, ref string[] names) =>
                    sapModel.LoadPatterns.GetNameList(ref numberNames, ref names),
                (ETABSv1.cSapModel sapModel, string name) =>
                {
                    ETABSv1.eLoadPatternType type = ETABSv1.eLoadPatternType.Dead;
                    sapModel.LoadPatterns.GetLoadType(name, ref type);
                    return type.ToString();
                });
            
            return patternResult;
        }

        public OperationResult DeleteLoadPatterns(IReadOnlyList<string> loadPatternNames)
        {
            var sapModelResult = EnsureEtabsSapModel();
            if (!sapModelResult.IsSuccess)
            {
                return OperationResult.Failure(sapModelResult.Message);
            }

            var result = Infrastructure.CSISapModel.LoadPatternService.CSISapModelLoadPatternService.DeleteLoadPatterns(
                sapModelResult.Data,
                loadPatternNames,
                (ETABSv1.cSapModel sapModel, string name) => sapModel.LoadPatterns.Delete(name));
            
            if (result.IsSuccess)
            {
                RefreshView(sapModelResult.Data);
            }

            return result;
        }

        public OperationResult<IReadOnlyList<CSISapModelFrameSectionDTO>> GetFrameSections()
        {
            var sapModelResult = EnsureEtabsSapModel();
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
                return OperationResult<IReadOnlyList<CSISapModelFrameSectionDTO>>.Failure("Failed to get frame section names from ETABS.");
            }

            var list = new List<CSISapModelFrameSectionDTO>();
            for (int i = 0; i < numberNames; i++)
            {
                ETABSv1.eFramePropType propType = ETABSv1.eFramePropType.I;
                sapModel.PropFrame.GetTypeOAPI(names[i], ref propType);

                FrameSectionShapeType shapeType = FrameSectionShapeType.Unknown;
                switch (propType)
                {
                    case ETABSv1.eFramePropType.I: shapeType = FrameSectionShapeType.I; break;
                    case ETABSv1.eFramePropType.Channel: shapeType = FrameSectionShapeType.Channel; break;
                    case ETABSv1.eFramePropType.T: shapeType = FrameSectionShapeType.T; break;
                    case ETABSv1.eFramePropType.Angle: shapeType = FrameSectionShapeType.Angle; break;
                    case ETABSv1.eFramePropType.DblAngle: shapeType = FrameSectionShapeType.DoubleAngle; break;
                    case ETABSv1.eFramePropType.Box: shapeType = FrameSectionShapeType.Tube; break;
                    case ETABSv1.eFramePropType.Pipe: shapeType = FrameSectionShapeType.Pipe; break;
                    case ETABSv1.eFramePropType.Rectangular: shapeType = FrameSectionShapeType.Rectangular; break;
                    case ETABSv1.eFramePropType.Circle: shapeType = FrameSectionShapeType.Circular; break;
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
            var sapModelResult = EnsureEtabsSapModel();
            if (!sapModelResult.IsSuccess)
            {
                return OperationResult<CSISapModelFrameSectionDetailDTO>.Failure(sapModelResult.Message);
            }

            var sapModel = sapModelResult.Data;
            ETABSv1.eFramePropType propType = ETABSv1.eFramePropType.I;
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
            double t3 = 0, t2 = 0, tf = 0, tw = 0, t2b = 0, tfb = 0, dis = 0;
            double area = 0, as2 = 0, as3 = 0, torsion = 0, i22 = 0, i33 = 0, s22 = 0, s33 = 0, z22 = 0, z33 = 0, r22 = 0, r33 = 0;

            switch (propType)
            {
                case ETABSv1.eFramePropType.Pipe:
                    detail.ShapeType = FrameSectionShapeType.Pipe;
                    sapModel.PropFrame.GetPipe(sectionName, ref fileName, ref matProp, ref t3, ref tw, ref color, ref notes, ref guid);
                    detail.Dimensions["Outside diameter ( t3 )"] = t3;
                    detail.Dimensions["Wall thickness ( tw )"] = tw;
                    detail.Color = color;
                    detail.Notes = notes;
                    break;
                case ETABSv1.eFramePropType.I:
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
                case ETABSv1.eFramePropType.Channel:
                    detail.ShapeType = FrameSectionShapeType.Channel;
                    sapModel.PropFrame.GetChannel(sectionName, ref fileName, ref matProp, ref t3, ref t2, ref tf, ref tw, ref color, ref notes, ref guid);
                    detail.Dimensions["Total depth ( t3 )"] = t3;
                    detail.Dimensions["Flange width ( t2 )"] = t2;
                    detail.Dimensions["Flange thickness ( tf )"] = tf;
                    detail.Dimensions["Web thickness ( tw )"] = tw;
                    detail.Color = color;
                    detail.Notes = notes;
                    break;
                case ETABSv1.eFramePropType.Angle:
                    detail.ShapeType = FrameSectionShapeType.Angle;
                    sapModel.PropFrame.GetAngle(sectionName, ref fileName, ref matProp, ref t3, ref t2, ref tf, ref tw, ref color, ref notes, ref guid);
                    detail.Dimensions["Total depth ( t3 )"] = t3;
                    detail.Dimensions["Flange width ( t2 )"] = t2;
                    detail.Dimensions["Flange thickness ( tf )"] = tf;
                    detail.Dimensions["Web thickness ( tw )"] = tw;
                    detail.Color = color;
                    detail.Notes = notes;
                    break;
                case ETABSv1.eFramePropType.DblAngle:
                    detail.ShapeType = FrameSectionShapeType.DoubleAngle;
                    sapModel.PropFrame.GetDblAngle(sectionName, ref fileName, ref matProp, ref t3, ref t2, ref tf, ref tw, ref dis, ref color, ref notes, ref guid);
                    detail.Dimensions["Total depth ( t3 )"] = t3;
                    detail.Dimensions["Flange width ( t2 )"] = t2;
                    detail.Dimensions["Flange thickness ( tf )"] = tf;
                    detail.Dimensions["Web thickness ( tw )"] = tw;
                    detail.Dimensions["Spacing ( dis )"] = dis;
                    detail.Color = color;
                    detail.Notes = notes;
                    break;
                case ETABSv1.eFramePropType.Rectangular:
                    detail.ShapeType = FrameSectionShapeType.Rectangular;
                    sapModel.PropFrame.GetRectangle(sectionName, ref fileName, ref matProp, ref t3, ref t2, ref color, ref notes, ref guid);
                    detail.Dimensions["Depth ( t3 )"] = t3;
                    detail.Dimensions["Width ( t2 )"] = t2;
                    detail.Color = color;
                    detail.Notes = notes;
                    break;
                case ETABSv1.eFramePropType.Circle:
                    detail.ShapeType = FrameSectionShapeType.Circular;
                    sapModel.PropFrame.GetCircle(sectionName, ref fileName, ref matProp, ref t3, ref color, ref notes, ref guid);
                    detail.Dimensions["Diameter ( t3 )"] = t3;
                    detail.Color = color;
                    detail.Notes = notes;
                    break;
                case ETABSv1.eFramePropType.Box:
                    detail.ShapeType = FrameSectionShapeType.Tube;
                    sapModel.PropFrame.GetTube(sectionName, ref fileName, ref matProp, ref t3, ref t2, ref tf, ref tw, ref color, ref notes, ref guid);
                    detail.Dimensions["Total depth ( t3 )"] = t3;
                    detail.Dimensions["Flange width ( t2 )"] = t2;
                    detail.Dimensions["Flange thickness ( tf )"] = tf;
                    detail.Dimensions["Web thickness ( tw )"] = tw;
                    detail.Color = color;
                    detail.Notes = notes;
                    break;
                case ETABSv1.eFramePropType.General:
                    detail.ShapeType = FrameSectionShapeType.General;
                    sapModel.PropFrame.GetGeneral(sectionName, ref fileName, ref matProp, ref t3, ref t2, ref area, ref as2, ref as3, ref torsion, ref i22, ref i33, ref s22, ref s33, ref z22, ref z33, ref r22, ref r33, ref color, ref notes, ref guid);
                    detail.Dimensions["Total depth ( t3 )"] = t3;
                    detail.Dimensions["Width ( t2 )"] = t2;
                    detail.Dimensions["Area"] = area;
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

        private static OperationResult RefreshView(ETABSv1.cSapModel sapModel)
        {
            int refreshResult = sapModel.View.RefreshView(0, false);
            if (refreshResult != 0)
            {
                return OperationResult.Failure($"ETABS model changed successfully, but View.RefreshView failed (return code {refreshResult}).");
            }

            return OperationResult.Success();
        }
    }
}

