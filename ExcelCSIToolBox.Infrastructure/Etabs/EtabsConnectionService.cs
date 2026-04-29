using System;
using System.Collections.Generic;
using ExcelCSIToolBox.Infrastructure.CSISapModel.Adapters;
using ExcelCSIToolBox.Core.Common.Results;
using ExcelCSIToolBox.Core.Abstractions.CSI;
using ExcelCSIToolBox.Core.Geometry;
using ExcelCSIToolBox.Core.Models.CSI;
using ExcelCSIToolBox.Data;
using ExcelCSIToolBox.Data.CSISapModel.FrameObject;
using ExcelCSIToolBox.Data.CSISapModel.PointObject;
using ExcelCSIToolBox.Data.DTOs.CSI;
using ExcelCSIToolBox.Data.Models;
using ExcelCSIToolBox.Infrastructure.CSISapModel;

namespace ExcelCSIToolBox.Infrastructure.Etabs
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

        public OperationResult ClearSelection()
        {
            var sapModelResult = EnsureEtabsSapModel();
            if (!sapModelResult.IsSuccess) return OperationResult.Failure(sapModelResult.Message);

            int ret = sapModelResult.Data.SelectObj.ClearSelection();
            return ret == 0 ? OperationResult.Success("Selection cleared.") : OperationResult.Failure($"Failed to clear selection (return code {ret}).");
        }

        public OperationResult AssignFrameSection(IReadOnlyList<string> frameNames, string sectionName)
        {
            var sapModelResult = EnsureEtabsSapModel();
            if (!sapModelResult.IsSuccess) return OperationResult.Failure(sapModelResult.Message);
            if (frameNames == null || frameNames.Count == 0) return OperationResult.Failure("At least one frame name is required.");
            if (string.IsNullOrWhiteSpace(sectionName)) return OperationResult.Failure("Section name is required.");

            var sapModel = sapModelResult.Data;
            if (!SectionNameExists(sapModel, sectionName)) return OperationResult.Failure($"Frame section '{sectionName}' does not exist.");

            int success = 0;
            var failures = new List<string>();
            foreach (string frameName in frameNames)
            {
                if (string.IsNullOrWhiteSpace(frameName)) continue;
                string p1 = string.Empty, p2 = string.Empty;
                if (sapModel.FrameObj.GetPoints(frameName, ref p1, ref p2) != 0)
                {
                    failures.Add($"{frameName}: not found");
                    continue;
                }

                int ret = sapModel.FrameObj.SetSection(frameName, sectionName, ETABSv1.eItemType.Objects, 0, 0);
                if (ret == 0) success++; else failures.Add($"{frameName}: return code {ret}");
            }

            RefreshView(sapModel);
            string msg = $"Assigned section '{sectionName}' to {success} frame(s).";
            if (failures.Count > 0) msg += " Failed: " + string.Join("; ", failures);
            return failures.Count == 0 ? OperationResult.Success(msg) : OperationResult.Failure(msg);
        }

        public OperationResult AssignFrameDistributedLoad(IReadOnlyList<string> frameNames, string loadPattern, int direction, double value1, double value2)
        {
            var sapModelResult = EnsureEtabsSapModel();
            if (!sapModelResult.IsSuccess) return OperationResult.Failure(sapModelResult.Message);
            if (frameNames == null || frameNames.Count == 0) return OperationResult.Failure("At least one frame name is required.");
            if (string.IsNullOrWhiteSpace(loadPattern)) return OperationResult.Failure("Load pattern is required.");

            var sapModel = sapModelResult.Data;
            int success = 0;
            var failures = new List<string>();
            foreach (string frameName in frameNames)
            {
                if (string.IsNullOrWhiteSpace(frameName)) continue;
                int ret = sapModel.FrameObj.SetLoadDistributed(frameName, loadPattern, 1, direction, 0, 1, value1, value2, "Global", true, true, ETABSv1.eItemType.Objects);
                if (ret == 0) success++; else failures.Add($"{frameName}: return code {ret}");
            }

            RefreshView(sapModel);
            string msg = $"Assigned distributed load '{loadPattern}' to {success} frame(s).";
            if (failures.Count > 0) msg += " Failed: " + string.Join("; ", failures);
            return failures.Count == 0 ? OperationResult.Success(msg) : OperationResult.Failure(msg);
        }

        public OperationResult AssignFramePointLoad(IReadOnlyList<string> frameNames, string loadPattern, int direction, double distance, double value)
        {
            var sapModelResult = EnsureEtabsSapModel();
            if (!sapModelResult.IsSuccess) return OperationResult.Failure(sapModelResult.Message);
            if (frameNames == null || frameNames.Count == 0) return OperationResult.Failure("At least one frame name is required.");
            if (string.IsNullOrWhiteSpace(loadPattern)) return OperationResult.Failure("Load pattern is required.");

            var sapModel = sapModelResult.Data;
            int success = 0;
            var failures = new List<string>();
            foreach (string frameName in frameNames)
            {
                if (string.IsNullOrWhiteSpace(frameName)) continue;
                int ret = sapModel.FrameObj.SetLoadPoint(frameName, loadPattern, 1, direction, distance, value, "Global", true, true, ETABSv1.eItemType.Objects);
                if (ret == 0) success++; else failures.Add($"{frameName}: return code {ret}");
            }

            RefreshView(sapModel);
            string msg = $"Assigned point load '{loadPattern}' to {success} frame(s).";
            if (failures.Count > 0) msg += " Failed: " + string.Join("; ", failures);
            return failures.Count == 0 ? OperationResult.Success(msg) : OperationResult.Failure(msg);
        }

        public OperationResult DeleteFrameObjects(IReadOnlyList<string> frameNames)
        {
            var sapModelResult = EnsureEtabsSapModel();
            if (!sapModelResult.IsSuccess) return OperationResult.Failure(sapModelResult.Message);
            if (frameNames == null || frameNames.Count == 0) return OperationResult.Failure("At least one frame name is required.");

            var sapModel = sapModelResult.Data;
            int success = 0;
            var failures = new List<string>();
            foreach (string frameName in frameNames)
            {
                if (string.IsNullOrWhiteSpace(frameName)) continue;
                int ret = sapModel.FrameObj.Delete(frameName, ETABSv1.eItemType.Objects);
                if (ret == 0) success++; else failures.Add($"{frameName}: return code {ret}");
            }

            RefreshView(sapModel);
            string msg = $"Deleted {success} frame object(s).";
            if (failures.Count > 0) msg += " Failed: " + string.Join("; ", failures);
            return failures.Count == 0 ? OperationResult.Success(msg) : OperationResult.Failure(msg);
        }

        public OperationResult RunAnalysis()
        {
            var sapModelResult = EnsureEtabsSapModel();
            if (!sapModelResult.IsSuccess) return OperationResult.Failure(sapModelResult.Message);
            int ret = sapModelResult.Data.Analyze.RunAnalysis();
            return ret == 0 ? OperationResult.Success("Analysis completed.") : OperationResult.Failure($"RunAnalysis failed (return code {ret}).");
        }

        public OperationResult SaveModel()
        {
            var sapModelResult = EnsureEtabsSapModel();
            if (!sapModelResult.IsSuccess) return OperationResult.Failure(sapModelResult.Message);
            string fileName = sapModelResult.Data.GetModelFilename(true);
            if (string.IsNullOrWhiteSpace(fileName)) return OperationResult.Failure("Model has no file path. Save is blocked.");
            int ret = sapModelResult.Data.File.Save(fileName);
            return ret == 0 ? OperationResult.Success("Model saved.") : OperationResult.Failure($"Save failed (return code {ret}).");
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

        public OperationResult<FrameAddBatchResultDto> AddFrameObjects(FrameAddBatchRequestDto request)
        {
            var sapModelResult = EnsureEtabsSapModel();
            if (!sapModelResult.IsSuccess)
            {
                return OperationResult<FrameAddBatchResultDto>.Failure("active CSI model is not available.");
            }

            return CSISapModelFrameObjectService.AddFrameObjects(
                request,
                ProductName,
                sapModelResult.Data,
                (ETABSv1.cSapModel sapModel, CSISapModelFrameByPointInput frameInput, ref string createdName, string sectionName, string userName) =>
                    sapModel.FrameObj.AddByPoint(
                        frameInput.Point1Name,
                        frameInput.Point2Name,
                        ref createdName,
                        sectionName,
                        userName),
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

        public OperationResult<IReadOnlyList<string>> GetPointNames()
        {
            var sapModelResult = EnsureEtabsSapModel();
            if (!sapModelResult.IsSuccess)
            {
                return OperationResult<IReadOnlyList<string>>.Failure(sapModelResult.Message);
            }

            return CSISapModelPointObjectService.GetNameList(
                ProductName,
                sapModelResult.Data,
                (ETABSv1.cSapModel sapModel, ref int numberNames, ref string[] names) =>
                    sapModel.PointObj.GetNameList(ref numberNames, ref names));
        }

        public OperationResult<PointObjectInfo> GetPointByName(string pointName)
        {
            return GetPointCoordinates(pointName);
        }

        public OperationResult<PointObjectInfo> GetPointCoordinates(string pointName)
        {
            var sapModelResult = EnsureEtabsSapModel();
            if (!sapModelResult.IsSuccess)
            {
                return OperationResult<PointObjectInfo>.Failure(sapModelResult.Message);
            }

            return CSISapModelPointObjectService.GetByName(
                ProductName,
                sapModelResult.Data,
                pointName,
                (ETABSv1.cSapModel sapModel, string name, ref double x, ref double y, ref double z) =>
                    sapModel.PointObj.GetCoordCartesian(name, ref x, ref y, ref z, "Global"),
                (ETABSv1.cSapModel sapModel, string name, ref bool selected) =>
                    sapModel.PointObj.GetSelected(name, ref selected));
        }

        public OperationResult<PointRestraintInfo> GetPointRestraint(string pointName)
        {
            var sapModelResult = EnsureEtabsSapModel();
            if (!sapModelResult.IsSuccess)
            {
                return OperationResult<PointRestraintInfo>.Failure(sapModelResult.Message);
            }

            return CSISapModelPointObjectService.GetRestraint(
                ProductName,
                sapModelResult.Data,
                pointName,
                (ETABSv1.cSapModel sapModel, string name, ref bool[] values) =>
                    sapModel.PointObj.GetRestraint(name, ref values));
        }

        public OperationResult<IReadOnlyList<PointLoadInfo>> GetPointLoadForces(string pointName)
        {
            var sapModelResult = EnsureEtabsSapModel();
            if (!sapModelResult.IsSuccess)
            {
                return OperationResult<IReadOnlyList<PointLoadInfo>>.Failure(sapModelResult.Message);
            }

            return CSISapModelPointObjectService.GetLoadForces(
                ProductName,
                sapModelResult.Data,
                pointName,
                (ETABSv1.cSapModel sapModel, string name, ref int numberItems, ref string[] pointNames, ref string[] loadPatterns, ref int[] caseSteps, ref string[] coordinateSystems, ref double[] f1, ref double[] f2, ref double[] f3, ref double[] m1, ref double[] m2, ref double[] m3) =>
                    sapModel.PointObj.GetLoadForce(name, ref numberItems, ref pointNames, ref loadPatterns, ref caseSteps, ref coordinateSystems, ref f1, ref f2, ref f3, ref m1, ref m2, ref m3, ETABSv1.eItemType.Objects));
        }

        public OperationResult<bool> GetPointSelected(string pointName)
        {
            var sapModelResult = EnsureEtabsSapModel();
            if (!sapModelResult.IsSuccess)
            {
                return OperationResult<bool>.Failure(sapModelResult.Message);
            }

            return CSISapModelPointObjectService.GetSelected(
                ProductName,
                sapModelResult.Data,
                pointName,
                (ETABSv1.cSapModel sapModel, string name, ref bool selected) =>
                    sapModel.PointObj.GetSelected(name, ref selected));
        }

        public OperationResult<string> GetPointGuid(string pointName)
        {
            var sapModelResult = EnsureEtabsSapModel();
            if (!sapModelResult.IsSuccess)
            {
                return OperationResult<string>.Failure(sapModelResult.Message);
            }

            return CSISapModelPointObjectService.GetGuid(
                ProductName,
                sapModelResult.Data,
                pointName,
                (ETABSv1.cSapModel sapModel, string name, ref string guid) =>
                    sapModel.PointObj.GetGUID(name, ref guid));
        }

        public OperationResult<PointGroupAssignmentInfo> GetPointGroupAssignments(string pointName)
        {
            var sapModelResult = EnsureEtabsSapModel();
            if (!sapModelResult.IsSuccess)
            {
                return OperationResult<PointGroupAssignmentInfo>.Failure(sapModelResult.Message);
            }

            return CSISapModelPointObjectService.GetGroupAssign(
                ProductName,
                sapModelResult.Data,
                pointName,
                (ETABSv1.cSapModel sapModel, string name, ref int numberItems, ref string[] groupNames) =>
                    sapModel.PointObj.GetGroupAssign(name, ref numberItems, ref groupNames));
        }

        public OperationResult<PointConnectivityInfo> GetPointConnectivity(string pointName)
        {
            var sapModelResult = EnsureEtabsSapModel();
            if (!sapModelResult.IsSuccess)
            {
                return OperationResult<PointConnectivityInfo>.Failure(sapModelResult.Message);
            }

            return CSISapModelPointObjectService.GetConnectivity(
                ProductName,
                sapModelResult.Data,
                pointName,
                (ETABSv1.cSapModel sapModel, string name, ref int numberItems, ref int[] objectTypes, ref string[] objectNames, ref int[] pointNumbers) =>
                    sapModel.PointObj.GetConnectivity(name, ref numberItems, ref objectTypes, ref objectNames, ref pointNumbers));
        }

        public OperationResult<PointSpringInfo> GetPointSpring(string pointName)
        {
            var sapModelResult = EnsureEtabsSapModel();
            if (!sapModelResult.IsSuccess)
            {
                return OperationResult<PointSpringInfo>.Failure(sapModelResult.Message);
            }

            return CSISapModelPointObjectService.GetSpring(
                ProductName,
                sapModelResult.Data,
                pointName,
                (ETABSv1.cSapModel sapModel, string name, ref double[] stiffness) =>
                    sapModel.PointObj.GetSpring(name, ref stiffness));
        }

        public OperationResult<PointMassInfo> GetPointMass(string pointName)
        {
            var sapModelResult = EnsureEtabsSapModel();
            if (!sapModelResult.IsSuccess)
            {
                return OperationResult<PointMassInfo>.Failure(sapModelResult.Message);
            }

            return CSISapModelPointObjectService.GetMass(
                ProductName,
                sapModelResult.Data,
                pointName,
                (ETABSv1.cSapModel sapModel, string name, ref double[] masses) =>
                    sapModel.PointObj.GetMass(name, ref masses));
        }

        public OperationResult<PointLocalAxesInfo> GetPointLocalAxes(string pointName)
        {
            var sapModelResult = EnsureEtabsSapModel();
            if (!sapModelResult.IsSuccess)
            {
                return OperationResult<PointLocalAxesInfo>.Failure(sapModelResult.Message);
            }

            return CSISapModelPointObjectService.GetLocalAxes(
                ProductName,
                sapModelResult.Data,
                pointName,
                (ETABSv1.cSapModel sapModel, string name, ref double a, ref double b, ref double c, ref bool advanced) =>
                    sapModel.PointObj.GetLocalAxes(name, ref a, ref b, ref c, ref advanced));
        }

        public OperationResult<PointDiaphragmInfo> GetPointDiaphragm(string pointName)
        {
            var sapModelResult = EnsureEtabsSapModel();
            if (!sapModelResult.IsSuccess)
            {
                return OperationResult<PointDiaphragmInfo>.Failure(sapModelResult.Message);
            }

            return CSISapModelPointObjectService.GetDiaphragm(
                ProductName,
                sapModelResult.Data,
                pointName,
                (ETABSv1.cSapModel sapModel, string name, ref int diaphragmOption, ref string diaphragmName) =>
                {
                    ETABSv1.eDiaphragmOption option = (ETABSv1.eDiaphragmOption)diaphragmOption;
                    int result = sapModel.PointObj.GetDiaphragm(name, ref option, ref diaphragmName);
                    diaphragmOption = (int)option;
                    return result;
                });
        }

        public OperationResult SetPointRestraint(IReadOnlyList<string> pointNames, IReadOnlyList<bool> restraints)
        {
            var sapModelResult = EnsureEtabsSapModel();
            if (!sapModelResult.IsSuccess)
            {
                return OperationResult.Failure(sapModelResult.Message);
            }

            return CSISapModelPointObjectService.SetRestraint(
                ProductName,
                sapModelResult.Data,
                pointNames,
                restraints,
                (ETABSv1.cSapModel sapModel, string name, ref bool[] values) =>
                    sapModel.PointObj.SetRestraint(name, ref values, ETABSv1.eItemType.Objects),
                RefreshView);
        }

        public OperationResult SetPointLoadForce(IReadOnlyList<string> pointNames, string loadPattern, IReadOnlyList<double> forceValues, bool replace, string coordinateSystem)
        {
            var sapModelResult = EnsureEtabsSapModel();
            if (!sapModelResult.IsSuccess)
            {
                return OperationResult.Failure(sapModelResult.Message);
            }

            return CSISapModelPointObjectService.SetLoadForce(
                ProductName,
                sapModelResult.Data,
                pointNames,
                loadPattern,
                forceValues,
                replace,
                coordinateSystem,
                (ETABSv1.cSapModel sapModel, string name, string pattern, ref double[] values, bool replaceExisting, string cSys) =>
                    sapModel.PointObj.SetLoadForce(name, pattern, ref values, replaceExisting, cSys, ETABSv1.eItemType.Objects),
                RefreshView);
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

        public OperationResult<IReadOnlyList<string>> GetFrameNames()
        {
            var sapModelResult = EnsureEtabsSapModel();
            if (!sapModelResult.IsSuccess)
            {
                return OperationResult<IReadOnlyList<string>>.Failure(sapModelResult.Message);
            }

            return CSISapModelFrameObjectService.GetNameList(
                ProductName,
                sapModelResult.Data,
                (ETABSv1.cSapModel sapModel, ref int numberNames, ref string[] names) =>
                    sapModel.FrameObj.GetNameList(ref numberNames, ref names));
        }

        public OperationResult<FrameObjectInfo> GetFrameByName(string frameName)
        {
            var sapModelResult = EnsureEtabsSapModel();
            if (!sapModelResult.IsSuccess)
            {
                return OperationResult<FrameObjectInfo>.Failure(sapModelResult.Message);
            }

            return CSISapModelFrameObjectService.GetByName(
                ProductName,
                sapModelResult.Data,
                frameName,
                (ETABSv1.cSapModel sapModel, string name, ref string pointI, ref string pointJ) =>
                    sapModel.FrameObj.GetPoints(name, ref pointI, ref pointJ),
                (ETABSv1.cSapModel sapModel, string name, ref string sectionName, ref string autoSelectList) =>
                    sapModel.FrameObj.GetSection(name, ref sectionName, ref autoSelectList),
                (ETABSv1.cSapModel sapModel, string name, ref bool selected) =>
                    sapModel.FrameObj.GetSelected(name, ref selected));
        }

        public OperationResult<FrameEndPointInfo> GetFramePoints(string frameName)
        {
            var sapModelResult = EnsureEtabsSapModel();
            if (!sapModelResult.IsSuccess)
            {
                return OperationResult<FrameEndPointInfo>.Failure(sapModelResult.Message);
            }

            return CSISapModelFrameObjectService.GetPoints(
                ProductName,
                sapModelResult.Data,
                frameName,
                (ETABSv1.cSapModel sapModel, string name, ref string pointI, ref string pointJ) =>
                    sapModel.FrameObj.GetPoints(name, ref pointI, ref pointJ));
        }

        public OperationResult<FrameSectionInfo> GetFrameSection(string frameName)
        {
            var sapModelResult = EnsureEtabsSapModel();
            if (!sapModelResult.IsSuccess)
            {
                return OperationResult<FrameSectionInfo>.Failure(sapModelResult.Message);
            }

            return CSISapModelFrameObjectService.GetSection(
                ProductName,
                sapModelResult.Data,
                frameName,
                (ETABSv1.cSapModel sapModel, string name, ref string sectionName, ref string autoSelectList) =>
                    sapModel.FrameObj.GetSection(name, ref sectionName, ref autoSelectList));
        }

        public OperationResult<IReadOnlyList<FrameLoadInfo>> GetFrameDistributedLoads(string frameName)
        {
            var sapModelResult = EnsureEtabsSapModel();
            if (!sapModelResult.IsSuccess)
            {
                return OperationResult<IReadOnlyList<FrameLoadInfo>>.Failure(sapModelResult.Message);
            }

            return CSISapModelFrameObjectService.GetDistributedLoads(
                ProductName,
                sapModelResult.Data,
                frameName,
                (ETABSv1.cSapModel sapModel, string name, ref int numberItems, ref string[] frameNames, ref string[] loadPatterns, ref int[] loadTypes, ref string[] coordinateSystems, ref int[] directions, ref double[] rd1, ref double[] rd2, ref double[] dist1, ref double[] dist2, ref double[] val1, ref double[] val2) =>
                    sapModel.FrameObj.GetLoadDistributed(name, ref numberItems, ref frameNames, ref loadPatterns, ref loadTypes, ref coordinateSystems, ref directions, ref rd1, ref rd2, ref dist1, ref dist2, ref val1, ref val2, ETABSv1.eItemType.Objects));
        }

        public OperationResult<IReadOnlyList<FrameLoadInfo>> GetFramePointLoads(string frameName)
        {
            var sapModelResult = EnsureEtabsSapModel();
            if (!sapModelResult.IsSuccess)
            {
                return OperationResult<IReadOnlyList<FrameLoadInfo>>.Failure(sapModelResult.Message);
            }

            return CSISapModelFrameObjectService.GetPointLoads(
                ProductName,
                sapModelResult.Data,
                frameName,
                (ETABSv1.cSapModel sapModel, string name, ref int numberItems, ref string[] frameNames, ref string[] loadPatterns, ref int[] loadTypes, ref string[] coordinateSystems, ref int[] directions, ref double[] relativeDistance, ref double[] distance, ref double[] value) =>
                    sapModel.FrameObj.GetLoadPoint(name, ref numberItems, ref frameNames, ref loadPatterns, ref loadTypes, ref coordinateSystems, ref directions, ref relativeDistance, ref distance, ref value, ETABSv1.eItemType.Objects));
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

        public OperationResult<IReadOnlyList<string>> GetShellNames()
        {
            var sapModelResult = EnsureEtabsSapModel();
            if (!sapModelResult.IsSuccess)
            {
                return OperationResult<IReadOnlyList<string>>.Failure(sapModelResult.Message);
            }

            return CSISapModelShellObjectService.GetNameList(
                sapModelResult.Data,
                "ETABS",
                (ETABSv1.cSapModel sapModel, ref int numberNames, ref string[] names) =>
                    sapModel.AreaObj.GetNameList(ref numberNames, ref names));
        }

        public OperationResult<CSISapModelShellObjectDTO> GetShellByName(string areaName)
        {
            var sapModelResult = EnsureEtabsSapModel();
            if (!sapModelResult.IsSuccess)
            {
                return OperationResult<CSISapModelShellObjectDTO>.Failure(sapModelResult.Message);
            }

            return CSISapModelShellObjectService.GetByName(
                sapModelResult.Data,
                "ETABS",
                areaName,
                (ETABSv1.cSapModel sapModel, string name, ref int numberPoints, ref string[] pointNames) =>
                    sapModel.AreaObj.GetPoints(name, ref numberPoints, ref pointNames),
                (ETABSv1.cSapModel sapModel, string name, ref string propertyName) =>
                    sapModel.AreaObj.GetProperty(name, ref propertyName),
                (ETABSv1.cSapModel sapModel, string name, ref bool selected) =>
                    sapModel.AreaObj.GetSelected(name, ref selected));
        }

        public OperationResult<IReadOnlyList<string>> GetShellPoints(string areaName)
        {
            var sapModelResult = EnsureEtabsSapModel();
            if (!sapModelResult.IsSuccess)
            {
                return OperationResult<IReadOnlyList<string>>.Failure(sapModelResult.Message);
            }

            return CSISapModelShellObjectService.GetPoints(
                sapModelResult.Data,
                "ETABS",
                areaName,
                (ETABSv1.cSapModel sapModel, string name, ref int numberPoints, ref string[] pointNames) =>
                    sapModel.AreaObj.GetPoints(name, ref numberPoints, ref pointNames));
        }

        public OperationResult<string> GetShellProperty(string areaName)
        {
            var sapModelResult = EnsureEtabsSapModel();
            if (!sapModelResult.IsSuccess)
            {
                return OperationResult<string>.Failure(sapModelResult.Message);
            }

            return CSISapModelShellObjectService.GetProperty(
                sapModelResult.Data,
                "ETABS",
                areaName,
                (ETABSv1.cSapModel sapModel, string name, ref string propertyName) =>
                    sapModel.AreaObj.GetProperty(name, ref propertyName));
        }

        public OperationResult<IReadOnlyList<string>> GetSelectedShells()
        {
            var sapModelResult = EnsureEtabsSapModel();
            if (!sapModelResult.IsSuccess)
            {
                return OperationResult<IReadOnlyList<string>>.Failure(sapModelResult.Message);
            }

            return CSISapModelShellObjectService.GetSelectedShells(
                sapModelResult.Data,
                "ETABS",
                (ETABSv1.cSapModel sapModel, ref int numberItems, ref int[] objectTypes, ref string[] objectNames) =>
                    sapModel.SelectObj.GetSelected(ref numberItems, ref objectTypes, ref objectNames));
        }

        public OperationResult<IReadOnlyList<CSISapModelShellLoadDTO>> GetShellUniformLoads(string areaName)
        {
            var sapModelResult = EnsureEtabsSapModel();
            if (!sapModelResult.IsSuccess)
            {
                return OperationResult<IReadOnlyList<CSISapModelShellLoadDTO>>.Failure(sapModelResult.Message);
            }

            return CSISapModelShellObjectService.GetUniformLoads(
                sapModelResult.Data,
                "ETABS",
                areaName,
                (ETABSv1.cSapModel sapModel, string name, ref int numberItems, ref string[] areaNames, ref string[] loadPatterns, ref string[] coordinateSystems, ref int[] directions, ref double[] values) =>
                    sapModel.AreaObj.GetLoadUniform(name, ref numberItems, ref areaNames, ref loadPatterns, ref coordinateSystems, ref directions, ref values, ETABSv1.eItemType.Objects));
        }

        public CsiWritePreview PreviewAddShellByPoint(IReadOnlyList<string> pointNames, string propertyName, string userName)
        {
            return CSISapModelShellObjectService.PreviewAddByPoint(pointNames, propertyName, userName);
        }

        public OperationResult<string> AddShellByPoint(IReadOnlyList<string> pointNames, string propertyName, string userName, bool confirmed)
        {
            var sapModelResult = EnsureEtabsSapModel();
            if (!sapModelResult.IsSuccess)
            {
                return OperationResult<string>.Failure(sapModelResult.Message);
            }

            return CSISapModelShellObjectService.AddByPoint(
                sapModelResult.Data,
                "ETABS",
                pointNames,
                propertyName,
                userName,
                confirmed,
                new CsiWriteGuard(),
                new CsiOperationLogger(),
                (ETABSv1.cSapModel sapModel, string pointName, ref double x, ref double y, ref double z) =>
                    sapModel.PointObj.GetCoordCartesian(pointName, ref x, ref y, ref z, "Global"),
                (ETABSv1.cSapModel sapModel, int numberPoints, ref string[] pointNamesArray, ref string areaName, string propName, string name) =>
                    sapModel.AreaObj.AddByPoint(numberPoints, ref pointNamesArray, ref areaName, propName, name),
                RefreshView);
        }

        public CsiWritePreview PreviewAddShellByCoord(IReadOnlyList<CSISapModelShellCoordinateInput> points, string propertyName, string userName, string coordinateSystem)
        {
            return CSISapModelShellObjectService.PreviewAddByCoord(points, propertyName, userName);
        }

        public OperationResult<string> AddShellByCoord(IReadOnlyList<CSISapModelShellCoordinateInput> points, string propertyName, string userName, string coordinateSystem, bool confirmed)
        {
            var sapModelResult = EnsureEtabsSapModel();
            if (!sapModelResult.IsSuccess)
            {
                return OperationResult<string>.Failure(sapModelResult.Message);
            }

            return CSISapModelShellObjectService.AddByCoord(
                sapModelResult.Data,
                "ETABS",
                points,
                propertyName,
                userName,
                coordinateSystem,
                confirmed,
                new CsiWriteGuard(),
                new CsiOperationLogger(),
                (ETABSv1.cSapModel sapModel, int numberPoints, ref double[] x, ref double[] y, ref double[] z, ref string areaName, string propName, string name, string cSys) =>
                    sapModel.AreaObj.AddByCoord(numberPoints, ref x, ref y, ref z, ref areaName, propName, name, cSys),
                RefreshView);
        }

        public CsiWritePreview PreviewAssignShellUniformLoad(IReadOnlyList<string> areaNames, string loadPattern, double value, int direction, bool replace, string coordinateSystem)
        {
            return CSISapModelShellObjectService.PreviewAssignUniformLoad(areaNames, loadPattern, value, direction, replace, coordinateSystem);
        }

        public OperationResult AssignShellUniformLoad(IReadOnlyList<string> areaNames, string loadPattern, double value, int direction, bool replace, string coordinateSystem, bool confirmed)
        {
            var sapModelResult = EnsureEtabsSapModel();
            if (!sapModelResult.IsSuccess)
            {
                return OperationResult.Failure(sapModelResult.Message);
            }

            return CSISapModelShellObjectService.AssignUniformLoad(
                sapModelResult.Data,
                "ETABS",
                areaNames,
                loadPattern,
                value,
                direction,
                replace,
                coordinateSystem,
                confirmed,
                new CsiWriteGuard(),
                new CsiOperationLogger(),
                (ETABSv1.cSapModel sapModel, string name, ref int numberPoints, ref string[] pointNames) =>
                    sapModel.AreaObj.GetPoints(name, ref numberPoints, ref pointNames),
                (ETABSv1.cSapModel sapModel, ref int numberNames, ref string[] names) =>
                    sapModel.LoadPatterns.GetNameList(ref numberNames, ref names),
                (ETABSv1.cSapModel sapModel, string name, string pattern, double loadValue, int loadDirection, bool loadReplace, string cSys) =>
                    sapModel.AreaObj.SetLoadUniform(name, pattern, loadValue, loadDirection, loadReplace, cSys, ETABSv1.eItemType.Objects),
                RefreshView);
        }

        public CsiWritePreview PreviewDeleteShells(IReadOnlyList<string> areaNames)
        {
            return CSISapModelShellObjectService.PreviewDeleteAreas(areaNames);
        }

        public OperationResult DeleteShells(IReadOnlyList<string> areaNames, bool confirmed)
        {
            var sapModelResult = EnsureEtabsSapModel();
            if (!sapModelResult.IsSuccess)
            {
                return OperationResult.Failure(sapModelResult.Message);
            }

            return CSISapModelShellObjectService.DeleteAreas(
                sapModelResult.Data,
                "ETABS",
                areaNames,
                confirmed,
                new CsiWriteGuard(),
                new CsiOperationLogger(),
                (ETABSv1.cSapModel sapModel, string name, ref int numberPoints, ref string[] pointNames) =>
                    sapModel.AreaObj.GetPoints(name, ref numberPoints, ref pointNames),
                (ETABSv1.cSapModel sapModel, string name) =>
                    sapModel.AreaObj.Delete(name, ETABSv1.eItemType.Objects),
                RefreshView);
        }

        public OperationResult<IReadOnlyList<ExcelCSIToolBox.Data.DTOs.CSI.CSISapModelLoadCombinationDTO>> GetLoadCombinations()
        {
            var sapModelResult = EnsureEtabsSapModel();
            if (!sapModelResult.IsSuccess)
            {
                var errorResult = OperationResult<IReadOnlyList<ExcelCSIToolBox.Data.DTOs.CSI.CSISapModelLoadCombinationDTO>>.Failure(sapModelResult.Message);
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

        public OperationResult<IReadOnlyList<ExcelCSIToolBox.Data.DTOs.CSI.LoadCombinationItemDTO>> GetLoadCombinationDetails(string combinationName)
        {
            var sapModelResult = EnsureEtabsSapModel();
            if (!sapModelResult.IsSuccess)
            {
                var errorResult = OperationResult<IReadOnlyList<ExcelCSIToolBox.Data.DTOs.CSI.LoadCombinationItemDTO>>.Failure(sapModelResult.Message);
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

        public OperationResult<IReadOnlyList<ExcelCSIToolBox.Data.DTOs.CSI.CSISapModelLoadPatternDTO>> GetLoadPatterns()
        {
            var sapModelResult = EnsureEtabsSapModel();
            if (!sapModelResult.IsSuccess)
            {
                var errorResult = OperationResult<IReadOnlyList<ExcelCSIToolBox.Data.DTOs.CSI.CSISapModelLoadPatternDTO>>.Failure(sapModelResult.Message);
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

        public OperationResult UpdateFrameSection(CSISapModelFrameSectionUpdateDTO input)
        {
            var sapModelResult = EnsureEtabsSapModel();
            if (!sapModelResult.IsSuccess) return OperationResult.Failure(sapModelResult.Message);

            var result = SetFrameSectionProperty(sapModelResult.Data, input.SectionName, input);
            if (!result.IsSuccess) return result;

            RefreshView(sapModelResult.Data);
            return OperationResult.Success("Frame section updated.");
        }

        public OperationResult RenameFrameSection(CSISapModelFrameSectionRenameDTO input)
        {
            var sapModelResult = EnsureEtabsSapModel();
            if (!sapModelResult.IsSuccess) return OperationResult.Failure(sapModelResult.Message);

            var sapModel = sapModelResult.Data;
            if (SectionNameExists(sapModel, input.SectionName))
            {
                return OperationResult.Failure($"Section '{input.SectionName}' already exists.");
            }

            var createResult = SetFrameSectionProperty(sapModel, input.SectionName, input);
            if (!createResult.IsSuccess) return createResult;

            int numberNames = 0;
            string[] frameNames = null;
            int listRet = sapModel.FrameObj.GetNameList(ref numberNames, ref frameNames);
            if (listRet != 0 || frameNames == null)
            {
                return OperationResult.Failure($"Created '{input.SectionName}', but failed to list frames for reassignment.");
            }

            int reassigned = 0;
            foreach (string frameName in frameNames)
            {
                string propName = string.Empty;
                string auto = string.Empty;
                if (sapModel.FrameObj.GetSection(frameName, ref propName, ref auto) == 0 &&
                    string.Equals(propName, input.OriginalName, StringComparison.Ordinal))
                {
                    int setRet = sapModel.FrameObj.SetSection(frameName, input.SectionName, ETABSv1.eItemType.Objects, 0, 0);
                    if (setRet != 0)
                    {
                        return OperationResult.Failure($"Created '{input.SectionName}', but failed to reassign frame '{frameName}'.");
                    }

                    reassigned++;
                }
            }

            int deleteRet = sapModel.PropFrame.Delete(input.OriginalName);
            RefreshView(sapModel);

            if (deleteRet != 0)
            {
                return OperationResult.Success($"Renamed section and reassigned {reassigned} frame(s). Old section could not be deleted automatically.");
            }

            return OperationResult.Success($"Renamed section and reassigned {reassigned} frame(s).");
        }

        private static OperationResult SetFrameSectionProperty(ETABSv1.cSapModel sapModel, string sectionName, CSISapModelFrameSectionUpdateDTO input)
        {
            if (string.IsNullOrWhiteSpace(sectionName)) return OperationResult.Failure("Section name is required.");
            if (string.IsNullOrWhiteSpace(input.MaterialName)) return OperationResult.Failure("Material name is required.");

            string notes = input.Notes ?? string.Empty;
            string guid = string.Empty;
            int ret;

            switch (input.ShapeType)
            {
                case FrameSectionShapeType.I:
                    ret = sapModel.PropFrame.SetISection(sectionName, input.MaterialName, Dim(input, "Total depth ( t3 )", "Depth ( t3 )"), Dim(input, "Top flange width ( t2 )", "Flange width ( t2 )"), Dim(input, "Top flange thickness ( tf )", "Flange thickness ( tf )"), Dim(input, "Web thickness ( tw )"), Dim(input, "Bottom flange width ( t2b )", "Top flange width ( t2 )", "Flange width ( t2 )"), Dim(input, "Bottom flange thickness ( tfb )", "Top flange thickness ( tf )", "Flange thickness ( tf )"), input.Color, notes, guid);
                    break;
                case FrameSectionShapeType.Channel:
                    ret = sapModel.PropFrame.SetChannel(sectionName, input.MaterialName, Dim(input, "Total depth ( t3 )", "Depth ( t3 )"), Dim(input, "Flange width ( t2 )", "Width ( t2 )"), Dim(input, "Flange thickness ( tf )"), Dim(input, "Web thickness ( tw )"), input.Color, notes, guid);
                    break;
                case FrameSectionShapeType.Angle:
                    ret = sapModel.PropFrame.SetAngle(sectionName, input.MaterialName, Dim(input, "Total depth ( t3 )", "Depth ( t3 )"), Dim(input, "Flange width ( t2 )", "Width ( t2 )"), Dim(input, "Flange thickness ( tf )"), Dim(input, "Web thickness ( tw )"), input.Color, notes, guid);
                    break;
                case FrameSectionShapeType.DoubleAngle:
                    ret = sapModel.PropFrame.SetDblAngle(sectionName, input.MaterialName, Dim(input, "Total depth ( t3 )", "Depth ( t3 )"), Dim(input, "Flange width ( t2 )", "Width ( t2 )"), Dim(input, "Flange thickness ( tf )"), Dim(input, "Web thickness ( tw )"), Dim(input, "Spacing ( dis )"), input.Color, notes, guid);
                    break;
                case FrameSectionShapeType.Tube:
                    ret = sapModel.PropFrame.SetTube_1(sectionName, input.MaterialName, Dim(input, "Total depth ( t3 )", "Depth ( t3 )"), Dim(input, "Flange width ( t2 )", "Width ( t2 )"), Dim(input, "Flange thickness ( tf )"), Dim(input, "Web thickness ( tw )"), 0.000000001, input.Color, notes, guid);
                    break;
                case FrameSectionShapeType.Pipe:
                    ret = sapModel.PropFrame.SetPipe(sectionName, input.MaterialName, Dim(input, "Outside diameter ( t3 )", "Diameter ( t3 )"), Dim(input, "Wall thickness ( tw )"), input.Color, notes, guid);
                    break;
                case FrameSectionShapeType.Rectangular:
                    ret = sapModel.PropFrame.SetRectangle(sectionName, input.MaterialName, Dim(input, "Depth ( t3 )", "Total depth ( t3 )"), Dim(input, "Width ( t2 )"), input.Color, notes, guid);
                    break;
                case FrameSectionShapeType.Circular:
                    ret = sapModel.PropFrame.SetCircle(sectionName, input.MaterialName, Dim(input, "Diameter ( t3 )", "Outside diameter ( t3 )"), input.Color, notes, guid);
                    break;
                default:
                    return OperationResult.Failure($"{input.ShapeType} editing is not supported yet.");
            }

            return ret == 0 ? OperationResult.Success() : OperationResult.Failure($"Failed to set frame section '{sectionName}' (return code {ret}).");
        }

        private static bool SectionNameExists(ETABSv1.cSapModel sapModel, string sectionName)
        {
            int numberNames = 0;
            string[] names = null;
            if (sapModel.PropFrame.GetNameList(ref numberNames, ref names) != 0 || names == null) return false;
            foreach (string name in names)
            {
                if (string.Equals(name, sectionName, StringComparison.Ordinal)) return true;
            }
            return false;
        }

        private static double Dim(CSISapModelFrameSectionUpdateDTO input, params string[] keys)
        {
            foreach (string key in keys)
            {
                if (input.Dimensions.TryGetValue(key, out double value)) return value;
            }
            return 0;
        }

        public OperationResult<CSISapModelStatisticsDTO> GetModelStatistics()
        {
            var sapModelResult = EnsureEtabsSapModel();
            if (!sapModelResult.IsSuccess) return OperationResult<CSISapModelStatisticsDTO>.Failure(sapModelResult.Message);
            var sapModel = sapModelResult.Data;

            var stats = new CSISapModelStatisticsDTO();

            try
            {
                int pointCount = 0;
                string[] pointNames = null;
                sapModel.PointObj.GetNameList(ref pointCount, ref pointNames);
                stats.PointCount = pointCount;

                int frameCount = 0;
                string[] frameNames = null;
                sapModel.FrameObj.GetNameList(ref frameCount, ref frameNames);
                stats.FrameCount = frameCount;

                int areaCount = 0;
                string[] areaNames = null;
                sapModel.AreaObj.GetNameList(ref areaCount, ref areaNames);
                stats.ShellCount = areaCount;

                int lpCount = 0;
                string[] lpNames = null;
                sapModel.LoadPatterns.GetNameList(ref lpCount, ref lpNames);
                stats.LoadPatternCount = lpCount;

                int comboCount = 0;
                string[] comboNames = null;
                sapModel.RespCombo.GetNameList(ref comboCount, ref comboNames);
                stats.LoadCombinationCount = comboCount;

                return OperationResult<CSISapModelStatisticsDTO>.Success(stats);
            }
            catch (Exception ex)
            {
                return OperationResult<CSISapModelStatisticsDTO>.Failure($"Failed to get model statistics: {ex.Message}");
            }
        }

        public OperationResult RefreshView(bool zoomAll = false)
        {
            var sapModelResult = EnsureEtabsSapModel();
            if (!sapModelResult.IsSuccess) return OperationResult.Failure(sapModelResult.Message);
            return RefreshView(sapModelResult.Data, zoomAll);
        }

        public OperationResult SetPresentUnits(int unitsCode)
        {
            var sapModelResult = EnsureEtabsSapModel();
            if (!sapModelResult.IsSuccess) return OperationResult.Failure(sapModelResult.Message);
            int ret = sapModelResult.Data.SetPresentUnits((ETABSv1.eUnits)unitsCode);
            return ret == 0 ? OperationResult.Success() : OperationResult.Failure($"Failed to set units (return code {ret}).");
        }

        private static OperationResult RefreshView(ETABSv1.cSapModel sapModel)
        {
            return RefreshView(sapModel, false);
        }

        private static OperationResult RefreshView(ETABSv1.cSapModel sapModel, bool zoomAll)
        {
            int refreshResult = sapModel.View.RefreshView(0, zoomAll);
            if (refreshResult != 0)
            {
                return OperationResult.Failure($"ETABS model changed successfully, but View.RefreshView failed (return code {refreshResult}).");
            }

            return OperationResult.Success();
        }

    }
}
