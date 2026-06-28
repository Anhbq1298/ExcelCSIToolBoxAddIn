using System;
using System.Collections.Generic;
using System.Globalization;
using ExcelCSIToolBox.Infrastructure.CSISapModel.Adapters;
using ExcelCSIToolBox.Core.Abstractions;
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
        private readonly IProgressReporter _progressReporter;

        public EtabsConnectionService()
            : this(CSISapModelConnectionAdapterFactory.CreateEtabs(), null)
        {
        }

        public EtabsConnectionService(ICsiModelAdapter modelAdapter, IProgressReporter progressReporter = null)
            : this(CSISapModelConnectionAdapterFactory.CreateEtabs(modelAdapter), progressReporter)
        {
        }

        private EtabsConnectionService(
            ICSISapModelConnectionAdapter<ETABSv1.cSapModel> connectionAdapter,
            IProgressReporter progressReporter)
        {
            _connectionAdapter = connectionAdapter ?? throw new ArgumentNullException(nameof(connectionAdapter));
            _progressReporter = progressReporter;
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
                RefreshView,
                _progressReporter);
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
                RefreshView,
                _progressReporter);
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
                RefreshView,
                _progressReporter);
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
                RefreshView,
                _progressReporter);
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
                RefreshView,
                _progressReporter);
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

        public OperationResult SetFrameReleases(IReadOnlyList<string> frameNames, IReadOnlyList<bool> startReleases, IReadOnlyList<bool> endReleases)
        {
            var sapModelResult = EnsureEtabsSapModel();
            if (!sapModelResult.IsSuccess)
            {
                return OperationResult.Failure("active CSI model is not available.");
            }

            if (frameNames == null || frameNames.Count == 0)
            {
                return OperationResult.Failure("At least one frame name is required.");
            }

            bool[] ii = ToReleaseArray(startReleases);
            bool[] jj = ToReleaseArray(endReleases);
            double[] startValues = new double[6];
            double[] endValues = new double[6];
            var failures = new List<string>();
            int success = 0;

            foreach (string frameName in frameNames)
            {
                if (string.IsNullOrWhiteSpace(frameName))
                {
                    continue;
                }

                int ret = sapModelResult.Data.FrameObj.SetReleases(frameName, ref ii, ref jj, ref startValues, ref endValues, ETABSv1.eItemType.Objects);
                if (ret == 0)
                {
                    success++;
                }
                else
                {
                    failures.Add($"{frameName}: return code {ret}");
                }
            }

            if (failures.Count > 0)
            {
                return OperationResult.Failure($"Set releases for {success} frame(s), failed {failures.Count}: {string.Join("; ", failures)}");
            }

            RefreshView(sapModelResult.Data);
            return OperationResult.Success($"Set releases for {success} frame(s).");
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
                inputs,
                _progressReporter);

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
                inputs,
                _progressReporter);

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
                inputs,
                _progressReporter);

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
                inputs,
                _progressReporter);

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
                inputs,
                _progressReporter);

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
                inputs,
                _progressReporter);

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
                inputs,
                _progressReporter);

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
                RefreshView,
                _progressReporter);
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

        public OperationResult<LoadCombinationMatrixDto> GetLoadCombinationMatrix()
        {
            var sapModelResult = EnsureEtabsSapModel();
            if (!sapModelResult.IsSuccess)
            {
                return OperationResult<LoadCombinationMatrixDto>.Failure(sapModelResult.Message);
            }

            return Infrastructure.CSISapModel.LoadCombinationService.CSISapModelLoadCombinationService.GetLoadCombinationMatrix(
                sapModelResult.Data,
                (ETABSv1.cSapModel sapModel, ref int numberNames, ref string[] names) =>
                    sapModel.LoadPatterns.GetNameList(ref numberNames, ref names),
                (ETABSv1.cSapModel sapModel, ref int numberNames, ref string[] names) =>
                    sapModel.RespCombo.GetNameList(ref numberNames, ref names),
                (ETABSv1.cSapModel sapModel, string name) =>
                {
                    int type = 0;
                    sapModel.RespCombo.GetTypeOAPI(name, ref type);
                    return type;
                },
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
                });
        }

        public OperationResult<IReadOnlyList<CSISapModelOutputCaseDTO>> GetAnalysisOutputCases()
        {
            var sapModelResult = EnsureEtabsSapModel();
            if (!sapModelResult.IsSuccess)
            {
                return OperationResult<IReadOnlyList<CSISapModelOutputCaseDTO>>.Failure(sapModelResult.Message);
            }

            try
            {
                var sapModel = sapModelResult.Data;
                var outputCases = new List<CSISapModelOutputCaseDTO>();
                var seenNames = new HashSet<string>(StringComparer.Ordinal);

                foreach (ETABSv1.eLoadCaseType caseType in Enum.GetValues(typeof(ETABSv1.eLoadCaseType)))
                {
                    int numberNames = 0;
                    string[] names = null;
                    int ret = sapModel.LoadCases.GetNameList(ref numberNames, ref names, caseType);
                    if (ret != 0 || names == null)
                    {
                        continue;
                    }

                    for (int i = 0; i < numberNames && i < names.Length; i++)
                    {
                        string name = names[i];
                        if (string.IsNullOrWhiteSpace(name) || !seenNames.Add(name))
                        {
                            continue;
                        }

                        string typeText = FormatLoadCaseType(caseType);
                        bool isSeismicWindOrRS = false;
                        if (caseType == ETABSv1.eLoadCaseType.ResponseSpectrum)
                        {
                            isSeismicWindOrRS = true;
                        }
                        else
                        {
                            var caseTypeVal = caseType;
                            int subType = 0;
                            var designType = ETABSv1.eLoadPatternType.Dead;
                            int designTypeOption = 0;
                            int auto = 0;
                            int typeRet = sapModel.LoadCases.GetTypeOAPI_1(
                                name,
                                ref caseTypeVal,
                                ref subType,
                                ref designType,
                                ref designTypeOption,
                                ref auto);
                            if (typeRet == 0)
                            {
                                isSeismicWindOrRS =
                                    designType == ETABSv1.eLoadPatternType.Quake ||
                                    designType == ETABSv1.eLoadPatternType.Wind ||
                                    designType == ETABSv1.eLoadPatternType.QuakeDrift ||
                                    designType == ETABSv1.eLoadPatternType.QuakeVerticalOnly;
                                if (isSeismicWindOrRS)
                                {
                                    typeText += $" ({designType})";
                                }
                            }
                        }

                        outputCases.Add(new CSISapModelOutputCaseDTO
                        {
                            Name = name,
                            Type = typeText,
                            IsLoadCombination = false,
                            IsSeismicWindOrResponseSpectrum = isSeismicWindOrRS
                        });
                    }
                }

                int numberCombos = 0;
                string[] comboNames = null;
                int comboRet = sapModel.RespCombo.GetNameList(ref numberCombos, ref comboNames);
                if (comboRet == 0 && comboNames != null)
                {
                    for (int i = 0; i < numberCombos && i < comboNames.Length; i++)
                    {
                        string name = comboNames[i];
                        if (string.IsNullOrWhiteSpace(name))
                        {
                            continue;
                        }

                        int type = 0;
                        sapModel.RespCombo.GetTypeOAPI(name, ref type);
                        outputCases.Add(new CSISapModelOutputCaseDTO
                        {
                            Name = name,
                            Type = FormatResponseCombinationType(type),
                            IsLoadCombination = true
                        });
                    }
                }

                return OperationResult<IReadOnlyList<CSISapModelOutputCaseDTO>>.Success(outputCases);
            }
            catch (Exception ex)
            {
                return OperationResult<IReadOnlyList<CSISapModelOutputCaseDTO>>.Failure($"Failed to load ETABS cases and combinations: {ex.Message}");
            }
        }

        public OperationResult<IReadOnlyList<CSISapModelOutputCaseDTO>> GetModalOutputCases()
        {
            var sapModelResult = EnsureEtabsSapModel();
            if (!sapModelResult.IsSuccess)
            {
                return OperationResult<IReadOnlyList<CSISapModelOutputCaseDTO>>.Failure(sapModelResult.Message);
            }

            try
            {
                var sapModel = sapModelResult.Data;
                var outputCases = new List<CSISapModelOutputCaseDTO>();
                var seenNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                AddModalCasesFromSpecificApi(
                    outputCases,
                    seenNames,
                    "Modal Eigen",
                    delegate(ref int numberNames, ref string[] names)
                    {
                        dynamic modalEigen = sapModel.LoadCases.ModalEigen;
                        return modalEigen.GetNameList(ref numberNames, ref names);
                    });

                AddModalCasesFromSpecificApi(
                    outputCases,
                    seenNames,
                    "Modal Ritz",
                    delegate(ref int numberNames, ref string[] names)
                    {
                        dynamic modalRitz = sapModel.LoadCases.ModalRitz;
                        return modalRitz.GetNameList(ref numberNames, ref names);
                    });

                if (outputCases.Count == 0)
                {
                    foreach (ETABSv1.eLoadCaseType caseType in Enum.GetValues(typeof(ETABSv1.eLoadCaseType)))
                    {
                        string typeText = FormatLoadCaseType(caseType);
                        if (typeText.IndexOf("Modal", StringComparison.OrdinalIgnoreCase) < 0 &&
                            typeText.IndexOf("Eigen", StringComparison.OrdinalIgnoreCase) < 0 &&
                            typeText.IndexOf("Ritz", StringComparison.OrdinalIgnoreCase) < 0)
                        {
                            continue;
                        }

                        int numberNames = 0;
                        string[] names = null;
                        int ret = sapModel.LoadCases.GetNameList(ref numberNames, ref names, caseType);
                        if (ret == 0)
                        {
                            AddModalCaseNames(outputCases, seenNames, names, numberNames, typeText);
                        }
                    }
                }

                return OperationResult<IReadOnlyList<CSISapModelOutputCaseDTO>>.Success(outputCases);
            }
            catch (Exception ex)
            {
                return OperationResult<IReadOnlyList<CSISapModelOutputCaseDTO>>.Failure($"Failed to load ETABS modal cases: {ex.Message}");
            }
        }

        public OperationResult<IReadOnlyList<CSISapModelOutputCaseDTO>> GetResponseSpectrumOutputCases()
        {
            var sapModelResult = EnsureEtabsSapModel();
            if (!sapModelResult.IsSuccess)
            {
                return OperationResult<IReadOnlyList<CSISapModelOutputCaseDTO>>.Failure(sapModelResult.Message);
            }

            try
            {
                var sapModel = sapModelResult.Data;
                var outputCases = new List<CSISapModelOutputCaseDTO>();
                var seenNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                int numberNames = 0;
                string[] names = null;
                int ret = sapModel.LoadCases.GetNameList(ref numberNames, ref names, ETABSv1.eLoadCaseType.ResponseSpectrum);
                if (ret != 0)
                {
                    return OperationResult<IReadOnlyList<CSISapModelOutputCaseDTO>>.Failure($"Failed to load ETABS response spectrum cases (return code {ret}).");
                }

                AddModalCaseNames(outputCases, seenNames, names, numberNames, "Response Spectrum");
                return OperationResult<IReadOnlyList<CSISapModelOutputCaseDTO>>.Success(outputCases);
            }
            catch (Exception ex)
            {
                return OperationResult<IReadOnlyList<CSISapModelOutputCaseDTO>>.Failure($"Failed to load ETABS response spectrum cases: {ex.Message}");
            }
        }

        public OperationResult<IReadOnlyList<CSISapModelBaseReactionRowDTO>> GetBaseReactions(IReadOnlyList<CSISapModelOutputCaseDTO> selectedOutputCases)
        {
            var sapModelResult = EnsureEtabsSapModel();
            if (!sapModelResult.IsSuccess)
            {
                return OperationResult<IReadOnlyList<CSISapModelBaseReactionRowDTO>>.Failure(sapModelResult.Message);
            }

            if (selectedOutputCases == null || selectedOutputCases.Count == 0)
            {
                return OperationResult<IReadOnlyList<CSISapModelBaseReactionRowDTO>>.Failure("Select at least one ETABS load case or load combination.");
            }

            try
            {
                var sapModel = sapModelResult.Data;
                int deselectRet = sapModel.Results.Setup.DeselectAllCasesAndCombosForOutput();
                if (deselectRet != 0)
                {
                    return OperationResult<IReadOnlyList<CSISapModelBaseReactionRowDTO>>.Failure($"Failed to clear ETABS output case selection (return code {deselectRet}).");
                }

                foreach (var outputCase in selectedOutputCases)
                {
                    if (outputCase == null || string.IsNullOrWhiteSpace(outputCase.Name))
                    {
                        continue;
                    }

                    int selectRet = outputCase.IsLoadCombination
                        ? sapModel.Results.Setup.SetComboSelectedForOutput(outputCase.Name, true)
                        : sapModel.Results.Setup.SetCaseSelectedForOutput(outputCase.Name, true);

                    if (selectRet != 0)
                    {
                        return OperationResult<IReadOnlyList<CSISapModelBaseReactionRowDTO>>.Failure($"Failed to select '{outputCase.Name}' for output (return code {selectRet}).");
                    }
                }

                string[] selectedCaseNamesForDisplay = CreateSelectedOutputCaseNameArray(selectedOutputCases, false);
                int caseDisplayRet = sapModel.DatabaseTables.SetLoadCasesSelectedForDisplay(ref selectedCaseNamesForDisplay);
                if (caseDisplayRet != 0)
                {
                    return OperationResult<IReadOnlyList<CSISapModelBaseReactionRowDTO>>.Failure($"Failed to select ETABS database table display load cases (return code {caseDisplayRet}).");
                }

                string[] selectedCombinationNamesForDisplay = CreateSelectedOutputCaseNameArray(selectedOutputCases, true);
                int combinationDisplayRet = sapModel.DatabaseTables.SetLoadCombinationsSelectedForDisplay(ref selectedCombinationNamesForDisplay);
                if (combinationDisplayRet != 0)
                {
                    return OperationResult<IReadOnlyList<CSISapModelBaseReactionRowDTO>>.Failure($"Failed to select ETABS database table display load combinations (return code {combinationDisplayRet}).");
                }

                string[] fieldKeyList = null;
                int tableVersion = 0;
                string[] fieldsKeysIncluded = null;
                int numberRecords = 0;
                string[] tableData = null;

                int ret = sapModel.DatabaseTables.GetTableForDisplayArray(
                    "Base Reactions",
                    ref fieldKeyList,
                    string.Empty,
                    ref tableVersion,
                    ref fieldsKeysIncluded,
                    ref numberRecords,
                    ref tableData);

                if (ret != 0)
                {
                    return OperationResult<IReadOnlyList<CSISapModelBaseReactionRowDTO>>.Failure($"Failed to read ETABS Base Reactions table (return code {ret}).");
                }

                string[] returnedFields = fieldsKeysIncluded != null && fieldsKeysIncluded.Length > 0
                    ? fieldsKeysIncluded
                    : fieldKeyList;
                IReadOnlyList<CSISapModelBaseReactionRowDTO> rows = FilterBaseReactionRows(
                    ParseBaseReactionRows(returnedFields, numberRecords, tableData),
                    selectedOutputCases);
                return OperationResult<IReadOnlyList<CSISapModelBaseReactionRowDTO>>.Success(rows);
            }
            catch (Exception ex)
            {
                return OperationResult<IReadOnlyList<CSISapModelBaseReactionRowDTO>>.Failure($"Failed to extract ETABS Base Reactions: {ex.Message}");
            }
        }

        public OperationResult<IReadOnlyList<CSISapModelModalMassParticipationRowDTO>> GetModalMassParticipationRatios(IReadOnlyList<CSISapModelOutputCaseDTO> selectedLoadCases)
        {
            var sapModelResult = EnsureEtabsSapModel();
            if (!sapModelResult.IsSuccess)
            {
                return OperationResult<IReadOnlyList<CSISapModelModalMassParticipationRowDTO>>.Failure(sapModelResult.Message);
            }

            if (selectedLoadCases == null || selectedLoadCases.Count == 0)
            {
                return OperationResult<IReadOnlyList<CSISapModelModalMassParticipationRowDTO>>.Failure("Select at least one ETABS modal load case.");
            }

            try
            {
                var sapModel = sapModelResult.Data;
                int deselectRet = sapModel.Results.Setup.DeselectAllCasesAndCombosForOutput();
                if (deselectRet != 0)
                {
                    return OperationResult<IReadOnlyList<CSISapModelModalMassParticipationRowDTO>>.Failure($"Failed to clear ETABS output case selection (return code {deselectRet}).");
                }

                foreach (var outputCase in selectedLoadCases)
                {
                    if (outputCase == null || string.IsNullOrWhiteSpace(outputCase.Name))
                    {
                        continue;
                    }

                    int selectRet = sapModel.Results.Setup.SetCaseSelectedForOutput(outputCase.Name, true);
                    if (selectRet != 0)
                    {
                        return OperationResult<IReadOnlyList<CSISapModelModalMassParticipationRowDTO>>.Failure($"Failed to select '{outputCase.Name}' for output (return code {selectRet}).");
                    }
                }

                string[] selectedCaseNamesForDisplay = CreateSelectedOutputCaseNameArray(selectedLoadCases, false);
                int caseDisplayRet = sapModel.DatabaseTables.SetLoadCasesSelectedForDisplay(ref selectedCaseNamesForDisplay);
                if (caseDisplayRet != 0)
                {
                    return OperationResult<IReadOnlyList<CSISapModelModalMassParticipationRowDTO>>.Failure($"Failed to select ETABS database table display load cases (return code {caseDisplayRet}).");
                }

                string[] selectedCombinationNamesForDisplay = new string[0];
                sapModel.DatabaseTables.SetLoadCombinationsSelectedForDisplay(ref selectedCombinationNamesForDisplay);

                string[] fieldKeyList = null;
                int tableVersion = 0;
                string[] fieldsKeysIncluded = null;
                int numberRecords = 0;
                string[] tableData = null;

                int ret = sapModel.DatabaseTables.GetTableForDisplayArray(
                    "Modal Participating Mass Ratios",
                    ref fieldKeyList,
                    string.Empty,
                    ref tableVersion,
                    ref fieldsKeysIncluded,
                    ref numberRecords,
                    ref tableData);

                if (ret != 0)
                {
                    return OperationResult<IReadOnlyList<CSISapModelModalMassParticipationRowDTO>>.Failure($"Failed to read ETABS Modal Participating Mass Ratios table (return code {ret}).");
                }

                string[] returnedFields = fieldsKeysIncluded != null && fieldsKeysIncluded.Length > 0
                    ? fieldsKeysIncluded
                    : fieldKeyList;
                IReadOnlyList<CSISapModelModalMassParticipationRowDTO> rows = FilterModalMassParticipationRows(
                    ParseModalMassParticipationRows(returnedFields, numberRecords, tableData),
                    selectedLoadCases);
                return OperationResult<IReadOnlyList<CSISapModelModalMassParticipationRowDTO>>.Success(rows);
            }
            catch (Exception ex)
            {
                return OperationResult<IReadOnlyList<CSISapModelModalMassParticipationRowDTO>>.Failure($"Failed to extract ETABS Modal Mass Participation Ratios: {ex.Message}");
            }
        }

        public OperationResult<IReadOnlyList<CSISapModelStoryForceRowDTO>> GetStoryForces(IReadOnlyList<CSISapModelOutputCaseDTO> selectedOutputCases)
        {
            var sapModelResult = EnsureEtabsSapModel();
            if (!sapModelResult.IsSuccess)
            {
                return OperationResult<IReadOnlyList<CSISapModelStoryForceRowDTO>>.Failure(sapModelResult.Message);
            }

            if (selectedOutputCases == null || selectedOutputCases.Count == 0)
            {
                return OperationResult<IReadOnlyList<CSISapModelStoryForceRowDTO>>.Failure("Select at least one ETABS load case or load combination.");
            }

            try
            {
                var sapModel = sapModelResult.Data;
                OperationResult selectResult = SelectOutputCasesForTables(sapModel, selectedOutputCases);
                if (!selectResult.IsSuccess)
                {
                    return OperationResult<IReadOnlyList<CSISapModelStoryForceRowDTO>>.Failure(selectResult.Message);
                }

                string[] fieldKeyList = null;
                int tableVersion = 0;
                string[] fieldsKeysIncluded = null;
                int numberRecords = 0;
                string[] tableData = null;

                int ret = sapModel.DatabaseTables.GetTableForDisplayArray(
                    "Story Forces",
                    ref fieldKeyList,
                    string.Empty,
                    ref tableVersion,
                    ref fieldsKeysIncluded,
                    ref numberRecords,
                    ref tableData);

                if (ret != 0)
                {
                    return OperationResult<IReadOnlyList<CSISapModelStoryForceRowDTO>>.Failure($"Failed to read ETABS Story Forces table (return code {ret}).");
                }

                string[] returnedFields = fieldsKeysIncluded != null && fieldsKeysIncluded.Length > 0
                    ? fieldsKeysIncluded
                    : fieldKeyList;
                IReadOnlyList<CSISapModelStoryForceRowDTO> rows = FilterStoryForceRows(
                    ParseStoryForceRows(returnedFields, numberRecords, tableData),
                    selectedOutputCases);
                return OperationResult<IReadOnlyList<CSISapModelStoryForceRowDTO>>.Success(rows);
            }
            catch (Exception ex)
            {
                return OperationResult<IReadOnlyList<CSISapModelStoryForceRowDTO>>.Failure($"Failed to extract ETABS Story Forces: {ex.Message}");
            }
        }

        public OperationResult<IReadOnlyList<CSISapModelStoryDisplacementRowDTO>> GetStoryDisplacements(IReadOnlyList<CSISapModelOutputCaseDTO> selectedOutputCases)
        {
            var sapModelResult = EnsureEtabsSapModel();
            if (!sapModelResult.IsSuccess)
            {
                return OperationResult<IReadOnlyList<CSISapModelStoryDisplacementRowDTO>>.Failure(sapModelResult.Message);
            }

            if (selectedOutputCases == null || selectedOutputCases.Count == 0)
            {
                return OperationResult<IReadOnlyList<CSISapModelStoryDisplacementRowDTO>>.Failure("Select at least one ETABS load case or load combination.");
            }

            try
            {
                var sapModel = sapModelResult.Data;
                OperationResult selectResult = SelectOutputCasesForTables(sapModel, selectedOutputCases);
                if (!selectResult.IsSuccess)
                {
                    return OperationResult<IReadOnlyList<CSISapModelStoryDisplacementRowDTO>>.Failure(selectResult.Message);
                }

                string[] fieldKeyList = null;
                int tableVersion = 0;
                string[] fieldsKeysIncluded = null;
                int numberRecords = 0;
                string[] tableData = null;

                int ret = sapModel.DatabaseTables.GetTableForDisplayArray(
                    "Story Displacements",
                    ref fieldKeyList,
                    string.Empty,
                    ref tableVersion,
                    ref fieldsKeysIncluded,
                    ref numberRecords,
                    ref tableData);

                if (ret != 0)
                {
                    return OperationResult<IReadOnlyList<CSISapModelStoryDisplacementRowDTO>>.Failure($"Failed to read ETABS Story Displacements table (return code {ret}).");
                }

                string[] returnedFields = fieldsKeysIncluded != null && fieldsKeysIncluded.Length > 0
                    ? fieldsKeysIncluded
                    : fieldKeyList;
                IReadOnlyList<CSISapModelStoryDisplacementRowDTO> rows = FilterStoryDisplacementRows(
                    ParseStoryDisplacementRows(returnedFields, numberRecords, tableData),
                    selectedOutputCases);
                return OperationResult<IReadOnlyList<CSISapModelStoryDisplacementRowDTO>>.Success(rows);
            }
            catch (Exception ex)
            {
                return OperationResult<IReadOnlyList<CSISapModelStoryDisplacementRowDTO>>.Failure($"Failed to extract ETABS Story Displacements: {ex.Message}");
            }
        }

        public OperationResult<CSISapModelDisplayTableDTO> GetStoryDrifts(IReadOnlyList<CSISapModelOutputCaseDTO> selectedOutputCases)
        {
            return GetStoryDisplayTable("Story Drifts", selectedOutputCases);
        }

        public OperationResult<CSISapModelDisplayTableDTO> GetStoryMaxOverAverageDisplacements(IReadOnlyList<CSISapModelOutputCaseDTO> selectedOutputCases)
        {
            return GetStoryDisplayTable("Story Max Over Avg Displacements", selectedOutputCases);
        }

        public OperationResult<CSISapModelDisplayTableDTO> GetStoryMaxOverAverageDrifts(IReadOnlyList<CSISapModelOutputCaseDTO> selectedOutputCases)
        {
            return GetStoryDisplayTable("Story Max Over Avg Drifts", selectedOutputCases);
        }

        public OperationResult<CSISapModelDisplayTableDTO> GetMassSummaryByStory()
        {
            var sapModelResult = EnsureEtabsSapModel();
            if (!sapModelResult.IsSuccess)
            {
                return OperationResult<CSISapModelDisplayTableDTO>.Failure(sapModelResult.Message);
            }

            try
            {
                string[] fieldKeyList = null;
                int tableVersion = 0;
                string[] fieldsKeysIncluded = null;
                int numberRecords = 0;
                string[] tableData = null;
                int ret = sapModelResult.Data.DatabaseTables.GetTableForDisplayArray(
                    "Mass Summary by Story",
                    ref fieldKeyList,
                    string.Empty,
                    ref tableVersion,
                    ref fieldsKeysIncluded,
                    ref numberRecords,
                    ref tableData);

                if (ret != 0)
                {
                    return OperationResult<CSISapModelDisplayTableDTO>.Failure($"Failed to read ETABS Mass Summary by Story table (return code {ret}).");
                }

                string[] returnedFields = fieldsKeysIncluded != null && fieldsKeysIncluded.Length > 0
                    ? fieldsKeysIncluded
                    : fieldKeyList;
                return OperationResult<CSISapModelDisplayTableDTO>.Success(ParseDisplayTable(returnedFields, numberRecords, tableData));
            }
            catch (Exception ex)
            {
                return OperationResult<CSISapModelDisplayTableDTO>.Failure($"Failed to extract ETABS Mass Summary by Story: {ex.Message}");
            }
        }

        public OperationResult<CSISapModelDisplayTableDTO> GetDisplayTable(string displayTableName)
        {
            return GetDisplayTable(displayTableName, null);
        }

        public OperationResult<CSISapModelDisplayTableDTO> GetDisplayTable(string displayTableName, IReadOnlyList<CSISapModelOutputCaseDTO> selectedOutputCases)
        {
            if (string.IsNullOrWhiteSpace(displayTableName))
            {
                return OperationResult<CSISapModelDisplayTableDTO>.Failure("ETABS database table name is required.");
            }

            var sapModelResult = EnsureEtabsSapModel();
            if (!sapModelResult.IsSuccess)
            {
                return OperationResult<CSISapModelDisplayTableDTO>.Failure(sapModelResult.Message);
            }

            try
            {
                ETABSv1.cSapModel sapModel = sapModelResult.Data;
                string normalized = NormalizeTableName(displayTableName);
                if (normalized == NormalizeTableName("Joint Displacements"))
                {
                    return GetGenericJointResultsFromOapi(sapModel, selectedOutputCases, displayTableName,
                        sapModel.Results.JointDispl, "U1", "U2", "U3", "R1", "R2", "R3");
                }
                if (normalized == NormalizeTableName("Joint Displacements - Absolute"))
                {
                    return GetGenericJointResultsFromOapi(sapModel, selectedOutputCases, displayTableName,
                        sapModel.Results.JointDisplAbs, "U1", "U2", "U3", "R1", "R2", "R3");
                }
                if (normalized == NormalizeTableName("Joint Drifts"))
                {
                    return GetJointDriftsFromOapi(sapModel, selectedOutputCases);
                }
                if (normalized == NormalizeTableName("Joint Reactions"))
                {
                    return GetGenericJointResultsFromOapi(sapModel, selectedOutputCases, displayTableName,
                        sapModel.Results.JointReact, "FX", "FY", "FZ", "MX", "MY", "MZ");
                }
                if (normalized == NormalizeTableName("Joint Velocities - Relative"))
                {
                    return GetGenericJointResultsFromOapi(sapModel, selectedOutputCases, displayTableName,
                        sapModel.Results.JointVel, "U1", "U2", "U3", "R1", "R2", "R3");
                }
                if (normalized == NormalizeTableName("Joint Velocities - Absolute"))
                {
                    return GetGenericJointResultsFromOapi(sapModel, selectedOutputCases, displayTableName,
                        sapModel.Results.JointVelAbs, "U1", "U2", "U3", "R1", "R2", "R3");
                }
                if (normalized == NormalizeTableName("Joint Accelerations - Relative"))
                {
                    return GetGenericJointResultsFromOapi(sapModel, selectedOutputCases, displayTableName,
                        sapModel.Results.JointAcc, "U1", "U2", "U3", "R1", "R2", "R3");
                }
                if (normalized == NormalizeTableName("Joint Accelerations - Absolute"))
                {
                    return GetGenericJointResultsFromOapi(sapModel, selectedOutputCases, displayTableName,
                        sapModel.Results.JointAccAbs, "U1", "U2", "U3", "R1", "R2", "R3");
                }

                if (selectedOutputCases != null && selectedOutputCases.Count > 0)
                {
                    OperationResult selectResult = IsCaseOnlyDisplayTable(displayTableName)
                        ? SelectLoadCasesOnlyForTables(sapModel, selectedOutputCases)
                        : SelectOutputCasesForTables(sapModel, selectedOutputCases);
                    if (!selectResult.IsSuccess)
                    {
                        return OperationResult<CSISapModelDisplayTableDTO>.Failure(selectResult.Message);
                    }
                }

                OperationResult<string> tableKeyResult = FindAvailableDisplayTableKey(sapModel, displayTableName);
                if (!tableKeyResult.IsSuccess)
                {
                    return OperationResult<CSISapModelDisplayTableDTO>.Failure(tableKeyResult.Message);
                }

                string[] fieldKeyList = null;
                int tableVersion = 0;
                string[] fieldsKeysIncluded = null;
                int numberRecords = 0;
                string[] tableData = null;

                int ret = sapModel.DatabaseTables.GetTableForDisplayArray(
                    tableKeyResult.Data,
                    ref fieldKeyList,
                    string.Empty,
                    ref tableVersion,
                    ref fieldsKeysIncluded,
                    ref numberRecords,
                    ref tableData);

                if (ret != 0)
                {
                    return OperationResult<CSISapModelDisplayTableDTO>.Failure($"Failed to read ETABS {displayTableName} table (return code {ret}).");
                }

                string[] returnedFields = fieldsKeysIncluded != null && fieldsKeysIncluded.Length > 0
                    ? fieldsKeysIncluded
                    : fieldKeyList;

                CSISapModelDisplayTableDTO table = ParseDisplayTable(returnedFields, numberRecords, tableData);
                if (selectedOutputCases != null && selectedOutputCases.Count > 0)
                {
                    table = FilterDisplayTableRows(table, selectedOutputCases, displayTableName);
                }

                var selectionInfo = GetActiveSelectionInfo(sapModel);
                if (selectionInfo.HasActiveSelection)
                {
                    if (IsJointOutputTable(displayTableName))
                    {
                        int jointColIndex = FindFieldIndex(
                            table.FieldKeys,
                            "Unique Name",
                            "UniqueName",
                            "Joint",
                            "Joint Name",
                            "JointName",
                            "Point",
                            "Point Name",
                            "PointName",
                            "Label",
                            "Label Name",
                            "LabelName");

                        if (jointColIndex >= 0)
                        {
                            var filteredRows = new List<object[]>();
                            foreach (object[] row in table.Rows)
                            {
                                string jointName = row != null && jointColIndex < row.Length
                                    ? Convert.ToString(row[jointColIndex], CultureInfo.InvariantCulture)
                                    : string.Empty;
                                if (!string.IsNullOrWhiteSpace(jointName) && selectionInfo.SelectedPoints.Contains(jointName.Trim()))
                                {
                                    filteredRows.Add(row);
                                }
                            }
                            table = new CSISapModelDisplayTableDTO { FieldKeys = table.FieldKeys, Rows = filteredRows };
                        }
                    }
                    else if (IsFrameOutputTable(displayTableName))
                    {
                        int frameColIndex = FindFieldIndex(
                            table.FieldKeys,
                            "Unique Name",
                            "UniqueName",
                            "Frame",
                            "Frame Name",
                            "FrameName",
                            "Element",
                            "Element Name",
                            "ElementName",
                            "Label");

                        if (frameColIndex >= 0)
                        {
                            var filteredRows = new List<object[]>();
                            foreach (object[] row in table.Rows)
                            {
                                string frameName = row != null && frameColIndex < row.Length
                                    ? Convert.ToString(row[frameColIndex], CultureInfo.InvariantCulture)
                                    : string.Empty;
                                if (!string.IsNullOrWhiteSpace(frameName) && selectionInfo.SelectedFrames.Contains(frameName.Trim()))
                                {
                                    filteredRows.Add(row);
                                }
                            }
                            table = new CSISapModelDisplayTableDTO { FieldKeys = table.FieldKeys, Rows = filteredRows };
                        }
                    }
                    else if (IsAreaOutputTable(displayTableName))
                    {
                        int areaColIndex = FindFieldIndex(
                            table.FieldKeys,
                            "Unique Name",
                            "UniqueName",
                            "Area",
                            "Area Name",
                            "AreaName",
                            "Element",
                            "Element Name",
                            "ElementName",
                            "Label");

                        if (areaColIndex >= 0)
                        {
                            var filteredRows = new List<object[]>();
                            foreach (object[] row in table.Rows)
                            {
                                string areaName = row != null && areaColIndex < row.Length
                                    ? Convert.ToString(row[areaColIndex], CultureInfo.InvariantCulture)
                                    : string.Empty;
                                if (!string.IsNullOrWhiteSpace(areaName) && selectionInfo.SelectedAreas.Contains(areaName.Trim()))
                                {
                                    filteredRows.Add(row);
                                }
                            }
                            table = new CSISapModelDisplayTableDTO { FieldKeys = table.FieldKeys, Rows = filteredRows };
                        }
                    }
                    else if (IsWallOutputTable(displayTableName))
                    {
                        int pierColIndex = FindFieldIndex(
                            table.FieldKeys,
                            "Pier",
                            "Pier Name",
                            "PierName",
                            "Label");

                        if (pierColIndex >= 0)
                        {
                            var filteredRows = new List<object[]>();
                            foreach (object[] row in table.Rows)
                            {
                                string pierName = row != null && pierColIndex < row.Length
                                    ? Convert.ToString(row[pierColIndex], CultureInfo.InvariantCulture)
                                    : string.Empty;
                                if (!string.IsNullOrWhiteSpace(pierName) && selectionInfo.SelectedPiers.Contains(pierName.Trim()))
                                {
                                    filteredRows.Add(row);
                                }
                            }
                            table = new CSISapModelDisplayTableDTO { FieldKeys = table.FieldKeys, Rows = filteredRows };
                        }
                    }
                }

                return OperationResult<CSISapModelDisplayTableDTO>.Success(table);
            }
            catch (Exception ex)
            {
                return OperationResult<CSISapModelDisplayTableDTO>.Failure($"Failed to extract ETABS {displayTableName}: {ex.Message}");
            }
        }

        private delegate int JointResultOapiDelegate(
            string name,
            ETABSv1.eItemTypeElm itemTypeElm,
            ref int numberResults,
            ref string[] obj,
            ref string[] elm,
            ref string[] loadCase,
            ref string[] stepType,
            ref double[] stepNum,
            ref double[] out1,
            ref double[] out2,
            ref double[] out3,
            ref double[] out4,
            ref double[] out5,
            ref double[] out6);

        private OperationResult<CSISapModelDisplayTableDTO> GetGenericJointResultsFromOapi(
            ETABSv1.cSapModel sapModel,
            IReadOnlyList<CSISapModelOutputCaseDTO> selectedOutputCases,
            string displayTableName,
            JointResultOapiDelegate oapiFunc,
            string col1, string col2, string col3, string col4, string col5, string col6)
        {
            var selectionInfo = GetActiveSelectionInfo(sapModel);
            ETABSv1.eItemTypeElm itemType = selectionInfo.HasActiveSelection
                ? ETABSv1.eItemTypeElm.SelectionElm
                : ETABSv1.eItemTypeElm.ObjectElm;

            if (selectedOutputCases != null && selectedOutputCases.Count > 0)
            {
                OperationResult selectResult = SelectOutputCasesForTables(sapModel, selectedOutputCases);
                if (!selectResult.IsSuccess)
                {
                    return OperationResult<CSISapModelDisplayTableDTO>.Failure(selectResult.Message);
                }
            }

            int numberResults = 0;
            string[] obj = null;
            string[] elm = null;
            string[] loadCase = null;
            string[] stepType = null;
            double[] stepNum = null;
            double[] out1 = null;
            double[] out2 = null;
            double[] out3 = null;
            double[] out4 = null;
            double[] out5 = null;
            double[] out6 = null;

            int ret = oapiFunc(
                "",
                itemType,
                ref numberResults,
                ref obj,
                ref elm,
                ref loadCase,
                ref stepType,
                ref stepNum,
                ref out1,
                ref out2,
                ref out3,
                ref out4,
                ref out5,
                ref out6);

            if (ret != 0)
            {
                return OperationResult<CSISapModelDisplayTableDTO>.Failure($"Failed to retrieve {displayTableName} from ETABS API (return code {ret}).");
            }

            var fieldKeys = new List<string>
            {
                "Unique Name",
                "Output Case",
                "Step Type",
                "Step Number",
                col1,
                col2,
                col3,
                col4,
                col5,
                col6
            };

            var rows = new List<object[]>();
            if (obj != null)
            {
                for (int i = 0; i < numberResults; i++)
                {
                    var row = new object[10];
                    row[0] = obj[i];
                    row[1] = loadCase != null && i < loadCase.Length ? loadCase[i] : "";
                    row[2] = stepType != null && i < stepType.Length ? stepType[i] : "";
                    row[3] = stepNum != null && i < stepNum.Length ? stepNum[i] : 0.0;
                    row[4] = out1 != null && i < out1.Length ? out1[i] : 0.0;
                    row[5] = out2 != null && i < out2.Length ? out2[i] : 0.0;
                    row[6] = out3 != null && i < out3.Length ? out3[i] : 0.0;
                    row[7] = out4 != null && i < out4.Length ? out4[i] : 0.0;
                    row[8] = out5 != null && i < out5.Length ? out5[i] : 0.0;
                    row[9] = out6 != null && i < out6.Length ? out6[i] : 0.0;
                    rows.Add(row);
                }
            }

            return OperationResult<CSISapModelDisplayTableDTO>.Success(new CSISapModelDisplayTableDTO
            {
                FieldKeys = fieldKeys,
                Rows = rows
            });
        }

        private OperationResult<CSISapModelDisplayTableDTO> GetJointDriftsFromOapi(
            ETABSv1.cSapModel sapModel,
            IReadOnlyList<CSISapModelOutputCaseDTO> selectedOutputCases)
        {
            var selectionInfo = GetActiveSelectionInfo(sapModel);
            bool hasSelection = selectionInfo.HasActiveSelection;
            var selectedJointNames = selectionInfo.SelectedPoints;

            if (selectedOutputCases != null && selectedOutputCases.Count > 0)
            {
                OperationResult selectResult = SelectOutputCasesForTables(sapModel, selectedOutputCases);
                if (!selectResult.IsSuccess)
                {
                    return OperationResult<CSISapModelDisplayTableDTO>.Failure(selectResult.Message);
                }
            }

            int numberResults = 0;
            string[] story = null;
            string[] label = null;
            string[] name = null;
            string[] loadCase = null;
            string[] stepType = null;
            double[] stepNum = null;
            double[] displacementX = null;
            double[] displacementY = null;
            double[] driftX = null;
            double[] driftY = null;

            int ret = sapModel.Results.JointDrifts(
                ref numberResults,
                ref story,
                ref label,
                ref name,
                ref loadCase,
                ref stepType,
                ref stepNum,
                ref displacementX,
                ref displacementY,
                ref driftX,
                ref driftY);

            if (ret != 0)
            {
                return OperationResult<CSISapModelDisplayTableDTO>.Failure($"Failed to retrieve Joint Drifts from ETABS API (return code {ret}).");
            }

            var fieldKeys = new List<string>
            {
                "Story",
                "Label",
                "Unique Name",
                "Output Case",
                "Step Type",
                "Step Number",
                "Displacement X",
                "Displacement Y",
                "Drift X",
                "Drift Y"
            };

            var rows = new List<object[]>();
            if (name != null)
            {
                for (int i = 0; i < numberResults; i++)
                {
                    string jointName = name[i];
                    string jointLabel = label != null && i < label.Length ? label[i] : "";
                    
                    if (!hasSelection || selectedJointNames.Contains(jointName.Trim()) || selectedJointNames.Contains(jointLabel.Trim()))
                    {
                        var row = new object[10];
                        row[0] = story != null && i < story.Length ? story[i] : "";
                        row[1] = jointLabel;
                        row[2] = jointName;
                        row[3] = loadCase != null && i < loadCase.Length ? loadCase[i] : "";
                        row[4] = stepType != null && i < stepType.Length ? stepType[i] : "";
                        row[5] = stepNum != null && i < stepNum.Length ? stepNum[i] : 0.0;
                        row[6] = displacementX != null && i < displacementX.Length ? displacementX[i] : 0.0;
                        row[7] = displacementY != null && i < displacementY.Length ? displacementY[i] : 0.0;
                        row[8] = driftX != null && i < driftX.Length ? driftX[i] : 0.0;
                        row[9] = driftY != null && i < driftY.Length ? driftY[i] : 0.0;
                        rows.Add(row);
                    }
                }
            }

            return OperationResult<CSISapModelDisplayTableDTO>.Success(new CSISapModelDisplayTableDTO
            {
                FieldKeys = fieldKeys,
                Rows = rows
            });
        }

        private OperationResult<CSISapModelDisplayTableDTO> GetObjectsAndElementsJoints(ETABSv1.cSapModel sapModel)
        {
            try
            {
                int numberNames = 0;
                string[] names = null;
                int ret = sapModel.PointObj.GetNameList(ref numberNames, ref names);
                if (ret != 0 || names == null)
                {
                    return OperationResult<CSISapModelDisplayTableDTO>.Failure("Failed to retrieve joint name list from ETABS.");
                }

                // Check for selected points
                var selectionInfo = GetActiveSelectionInfo(sapModel);
                var selectedNames = selectionInfo.SelectedPoints;
                bool hasSelection = selectionInfo.HasActiveSelection;

                var fieldKeys = new[] { "Joint", "Label", "Unique Name", "Story", "X-Coord", "Y-Coord", "Z-Coord" };
                var rows = new List<object[]>();

                foreach (var name in names)
                {
                    if (string.IsNullOrWhiteSpace(name)) continue;
                    if (hasSelection && !selectedNames.Contains(name.Trim())) continue;

                    string label = string.Empty;
                    string story = string.Empty;
                    sapModel.PointObj.GetLabelFromName(name, ref label, ref story);

                    double x = 0, y = 0, z = 0;
                    sapModel.PointObj.GetCoordCartesian(name, ref x, ref y, ref z, "Global");

                    rows.Add(new object[] { label, label, name, story, x, y, z });
                }

                return OperationResult<CSISapModelDisplayTableDTO>.Success(new CSISapModelDisplayTableDTO
                {
                    FieldKeys = fieldKeys,
                    Rows = rows
                });
            }
            catch (Exception ex)
            {
                return OperationResult<CSISapModelDisplayTableDTO>.Failure($"Failed to retrieve Objects and Elements - Joints: {ex.Message}");
            }
        }

        private OperationResult<CSISapModelDisplayTableDTO> GetObjectsAndElementsFrames(ETABSv1.cSapModel sapModel)
        {
            try
            {
                int numberNames = 0;
                string[] names = null;
                int ret = sapModel.FrameObj.GetNameList(ref numberNames, ref names);
                if (ret != 0 || names == null)
                {
                    return OperationResult<CSISapModelDisplayTableDTO>.Failure("Failed to retrieve frame name list from ETABS.");
                }

                // Check for selected frames
                var selectionInfo = GetActiveSelectionInfo(sapModel);
                var selectedNames = selectionInfo.SelectedFrames;
                bool hasSelection = selectionInfo.HasActiveSelection;

                var fieldKeys = new[] { "Frame", "Label", "Unique Name", "Story", "PointI", "PointJ", "Section", "Material" };
                var rows = new List<object[]>();

                foreach (var name in names)
                {
                    if (string.IsNullOrWhiteSpace(name)) continue;
                    if (hasSelection && !selectedNames.Contains(name.Trim())) continue;

                    string label = string.Empty;
                    string story = string.Empty;
                    sapModel.FrameObj.GetLabelFromName(name, ref label, ref story);

                    string pointI = string.Empty;
                    string pointJ = string.Empty;
                    sapModel.FrameObj.GetPoints(name, ref pointI, ref pointJ);

                    string section = string.Empty;
                    string sAuto = string.Empty;
                    sapModel.FrameObj.GetSection(name, ref section, ref sAuto);

                    string material = string.Empty;
                    if (!string.IsNullOrWhiteSpace(section))
                    {
                        sapModel.PropFrame.GetMaterial(section, ref material);
                    }

                    rows.Add(new object[] { label, label, name, story, pointI, pointJ, section, material });
                }

                return OperationResult<CSISapModelDisplayTableDTO>.Success(new CSISapModelDisplayTableDTO
                {
                    FieldKeys = fieldKeys,
                    Rows = rows
                });
            }
            catch (Exception ex)
            {
                return OperationResult<CSISapModelDisplayTableDTO>.Failure($"Failed to retrieve Objects and Elements - Frames: {ex.Message}");
            }
        }

        private OperationResult<CSISapModelDisplayTableDTO> GetObjectsAndElementsAreas(ETABSv1.cSapModel sapModel)
        {
            try
            {
                int numberNames = 0;
                string[] names = null;
                int ret = sapModel.AreaObj.GetNameList(ref numberNames, ref names);
                if (ret != 0 || names == null)
                {
                    return OperationResult<CSISapModelDisplayTableDTO>.Failure("Failed to retrieve area name list from ETABS.");
                }

                // Check for selected shells
                var selectionInfo = GetActiveSelectionInfo(sapModel);
                var selectedNames = selectionInfo.SelectedAreas;
                bool hasSelection = selectionInfo.HasActiveSelection;

                var fieldKeys = new[] { "Area", "Label", "Unique Name", "Story", "Section", "Material" };
                var rows = new List<object[]>();

                foreach (var name in names)
                {
                    if (string.IsNullOrWhiteSpace(name)) continue;
                    if (hasSelection && !selectedNames.Contains(name.Trim())) continue;

                    string label = string.Empty;
                    string story = string.Empty;
                    sapModel.AreaObj.GetLabelFromName(name, ref label, ref story);

                    string section = string.Empty;
                    sapModel.AreaObj.GetProperty(name, ref section);

                    string material = string.Empty;
                    if (!string.IsNullOrWhiteSpace(section))
                    {
                        // Try GetSlab
                        ETABSv1.eSlabType slabType = ETABSv1.eSlabType.Slab;
                        ETABSv1.eShellType shellType = ETABSv1.eShellType.ShellThin;
                        double thickness = 0;
                        int color = 0;
                        string notes = string.Empty;
                        string guid = string.Empty;
                        int sRet = sapModel.PropArea.GetSlab(section, ref slabType, ref shellType, ref material, ref thickness, ref color, ref notes, ref guid);
                        if (sRet != 0 || string.IsNullOrWhiteSpace(material))
                        {
                            // Try GetWall
                            ETABSv1.eWallPropType wallType = ETABSv1.eWallPropType.Specified;
                            int wRet = sapModel.PropArea.GetWall(section, ref wallType, ref shellType, ref material, ref thickness, ref color, ref notes, ref guid);
                            if (wRet != 0 || string.IsNullOrWhiteSpace(material))
                            {
                                // Try GetDeck
                                ETABSv1.eDeckType deckType = ETABSv1.eDeckType.Filled;
                                sapModel.PropArea.GetDeck(section, ref deckType, ref shellType, ref material, ref thickness, ref color, ref notes, ref guid);
                            }
                        }
                    }

                    rows.Add(new object[] { label, label, name, story, section, material });
                }

                return OperationResult<CSISapModelDisplayTableDTO>.Success(new CSISapModelDisplayTableDTO
                {
                    FieldKeys = fieldKeys,
                    Rows = rows
                });
            }
            catch (Exception ex)
            {
                return OperationResult<CSISapModelDisplayTableDTO>.Failure($"Failed to retrieve Objects and Elements - Areas: {ex.Message}");
            }
        }
        private struct ActiveSelectionInfo
        {
            public bool HasActiveSelection;
            public HashSet<string> SelectedPoints;
            public HashSet<string> SelectedFrames;
            public HashSet<string> SelectedAreas;
            public HashSet<string> SelectedPiers;
        }

        private ActiveSelectionInfo GetActiveSelectionInfo(ETABSv1.cSapModel sapModel)
        {
            var info = new ActiveSelectionInfo
            {
                HasActiveSelection = false,
                SelectedPoints = new HashSet<string>(StringComparer.OrdinalIgnoreCase),
                SelectedFrames = new HashSet<string>(StringComparer.OrdinalIgnoreCase),
                SelectedAreas = new HashSet<string>(StringComparer.OrdinalIgnoreCase),
                SelectedPiers = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            };

            try
            {
                int numberItems = 0;
                int[] objectTypes = null;
                string[] objectNames = null;
                int ret = sapModel.SelectObj.GetSelected(ref numberItems, ref objectTypes, ref objectNames);
                if (ret == 0 && numberItems > 0 && objectTypes != null && objectNames != null)
                {
                    info.HasActiveSelection = true;
                    for (int i = 0; i < numberItems; i++)
                    {
                        if (i >= objectTypes.Length || i >= objectNames.Length) continue;
                        string name = objectNames[i];
                        if (string.IsNullOrWhiteSpace(name)) continue;
                        name = name.Trim();

                        if (objectTypes[i] == ExcelCSIToolBox.Data.CSISapModelObjectTypeIds.Point)
                        {
                            info.SelectedPoints.Add(name);
                        }
                        else if (objectTypes[i] == ExcelCSIToolBox.Data.CSISapModelObjectTypeIds.Frame)
                        {
                            info.SelectedFrames.Add(name);
                            string pierName = string.Empty;
                            if (sapModel.FrameObj.GetPier(name, ref pierName) == 0 && !string.IsNullOrWhiteSpace(pierName))
                            {
                                info.SelectedPiers.Add(pierName.Trim());
                            }
                        }
                        else if (objectTypes[i] == ExcelCSIToolBox.Data.CSISapModelObjectTypeIds.Shell)
                        {
                            info.SelectedAreas.Add(name);
                            string pierName = string.Empty;
                            if (sapModel.AreaObj.GetPier(name, ref pierName) == 0 && !string.IsNullOrWhiteSpace(pierName))
                            {
                                info.SelectedPiers.Add(pierName.Trim());
                            }
                        }
                    }
                }
            }
            catch
            {
                // Ignore any error and return empty/no active selection info
            }

            return info;
        }

        private static bool IsJointOutputTable(string displayTableName)
        {
            string normalized = NormalizeTableName(displayTableName);
            return normalized == NormalizeTableName("Joint Displacements") ||
                   normalized == NormalizeTableName("Joint Displacements - Absolute") ||
                   normalized == NormalizeTableName("Joint Drifts") ||
                   normalized == NormalizeTableName("Joint Reactions") ||
                   normalized == NormalizeTableName("Joint Design Reactions") ||
                   normalized == NormalizeTableName("Joint Velocities - Relative") ||
                   normalized == NormalizeTableName("Joint Velocities - Absolute") ||
                   normalized == NormalizeTableName("Joint Accelerations - Relative") ||
                   normalized == NormalizeTableName("Joint Accelerations - Absolute") ||
                   normalized == NormalizeTableName("Assembled Joint Masses") ||
                   normalized == NormalizeTableName("Objects and Elements - Joints");
        }

        private static bool IsFrameOutputTable(string displayTableName)
        {
            string normalized = NormalizeTableName(displayTableName);
            return normalized == NormalizeTableName("Element Forces - Columns") ||
                   normalized == NormalizeTableName("Element Forces - Beams") ||
                   normalized == NormalizeTableName("Element Forces - Braces") ||
                   normalized == NormalizeTableName("Element Joint Forces - Frame");
        }

        private static bool IsAreaOutputTable(string displayTableName)
        {
            string normalized = NormalizeTableName(displayTableName);
            return normalized == NormalizeTableName("Element Forces - Area Shells") ||
                   normalized == NormalizeTableName("Element Stresses - Area Shells") ||
                   normalized == NormalizeTableName("Element Strains - Area Shells") ||
                   normalized == NormalizeTableName("Element Joint Forces - Shells");
        }

        private static bool IsWallOutputTable(string displayTableName)
        {
            string normalized = NormalizeTableName(displayTableName);
            return normalized == NormalizeTableName("Pier Forces");
        }

        private static OperationResult<string> FindAvailableDisplayTableKey(ETABSv1.cSapModel sapModel, string displayTableName)
        {
            int numberTables = 0;
            string[] tableKeys = null;
            string[] tableNames = null;
            int[] importTypes = null;

            int ret = sapModel.DatabaseTables.GetAvailableTables(ref numberTables, ref tableKeys, ref tableNames, ref importTypes);
            if (ret != 0)
            {
                return OperationResult<string>.Failure($"Failed to read ETABS available database tables (return code {ret}).");
            }

            string requested = NormalizeTableName(displayTableName);
            for (int index = 0; index < numberTables; index++)
            {
                string tableKey = tableKeys != null && index < tableKeys.Length ? tableKeys[index] : string.Empty;
                string tableName = tableNames != null && index < tableNames.Length ? tableNames[index] : string.Empty;

                if (string.Equals(NormalizeTableName(tableName), requested, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(NormalizeTableName(tableKey), requested, StringComparison.OrdinalIgnoreCase))
                {
                    return OperationResult<string>.Success(string.IsNullOrWhiteSpace(tableKey) ? tableName : tableKey);
                }
            }

            return OperationResult<string>.Failure("Table not available in this ETABS version or current analysis result.");
        }

        private static string NormalizeTableName(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return string.Empty;
            }

            var chars = new List<char>();
            foreach (char ch in value)
            {
                if (char.IsLetterOrDigit(ch))
                {
                    chars.Add(char.ToUpperInvariant(ch));
                }
            }

            return new string(chars.ToArray());
        }

        private static bool IsCaseOnlyDisplayTable(string displayTableName)
        {
            string normalized = NormalizeTableName(displayTableName);
            return normalized == NormalizeTableName("Modal Periods And Frequencies") ||
                   normalized == NormalizeTableName("Modal Participating Mass Ratios") ||
                   normalized == NormalizeTableName("Modal Load Participation Ratios") ||
                   normalized == NormalizeTableName("Modal Participation Factors") ||
                   normalized == NormalizeTableName("Modal Direction Factors") ||
                   normalized == NormalizeTableName("Response Spectrum Modal Info");
        }

        private static bool IsResponseSpectrumModalInfoTable(string displayTableName)
        {
            return NormalizeTableName(displayTableName) == NormalizeTableName("Response Spectrum Modal Info");
        }

        private OperationResult<CSISapModelDisplayTableDTO> GetStoryDisplayTable(
            string tableKey,
            IReadOnlyList<CSISapModelOutputCaseDTO> selectedOutputCases)
        {
            var sapModelResult = EnsureEtabsSapModel();
            if (!sapModelResult.IsSuccess)
            {
                return OperationResult<CSISapModelDisplayTableDTO>.Failure(sapModelResult.Message);
            }

            if (selectedOutputCases == null || selectedOutputCases.Count == 0)
            {
                return OperationResult<CSISapModelDisplayTableDTO>.Failure("Select at least one ETABS load case or load combination.");
            }

            try
            {
                var sapModel = sapModelResult.Data;
                OperationResult selectResult = SelectOutputCasesForTables(sapModel, selectedOutputCases);
                if (!selectResult.IsSuccess)
                {
                    return OperationResult<CSISapModelDisplayTableDTO>.Failure(selectResult.Message);
                }

                string[] fieldKeyList = null;
                int tableVersion = 0;
                string[] fieldsKeysIncluded = null;
                int numberRecords = 0;
                string[] tableData = null;
                int ret = sapModel.DatabaseTables.GetTableForDisplayArray(
                    tableKey,
                    ref fieldKeyList,
                    string.Empty,
                    ref tableVersion,
                    ref fieldsKeysIncluded,
                    ref numberRecords,
                    ref tableData);

                if (ret != 0)
                {
                    return OperationResult<CSISapModelDisplayTableDTO>.Failure($"Failed to read ETABS {tableKey} table (return code {ret}).");
                }

                string[] returnedFields = fieldsKeysIncluded != null && fieldsKeysIncluded.Length > 0
                    ? fieldsKeysIncluded
                    : fieldKeyList;
                return OperationResult<CSISapModelDisplayTableDTO>.Success(
                    FilterDisplayTableRows(ParseDisplayTable(returnedFields, numberRecords, tableData), selectedOutputCases));
            }
            catch (Exception ex)
            {
                return OperationResult<CSISapModelDisplayTableDTO>.Failure($"Failed to extract ETABS {tableKey}: {ex.Message}");
            }
        }

        public OperationResult<IReadOnlyList<string>> GetLoadPatternNames()
        {
            var sapModelResult = EnsureEtabsSapModel();
            if (!sapModelResult.IsSuccess)
            {
                return OperationResult<IReadOnlyList<string>>.Failure(sapModelResult.Message);
            }

            return Infrastructure.CSISapModel.LoadCombinationService.CSISapModelLoadCombinationService.GetLoadPatternNames(
                sapModelResult.Data,
                (ETABSv1.cSapModel sapModel, ref int numberNames, ref string[] names) =>
                    sapModel.LoadPatterns.GetNameList(ref numberNames, ref names));
        }

        public OperationResult<LoadCombinationApplyResultDto> ApplyLoadCombinationMatrix(IReadOnlyList<LoadCombinationMatrixRowDto> rows)
        {
            var sapModelResult = EnsureEtabsSapModel();
            if (!sapModelResult.IsSuccess)
            {
                return OperationResult<LoadCombinationApplyResultDto>.Failure(sapModelResult.Message);
            }

            var result = Infrastructure.CSISapModel.LoadCombinationService.CSISapModelLoadCombinationService.ApplyLoadCombinationMatrix(
                sapModelResult.Data,
                rows,
                (ETABSv1.cSapModel sapModel, ref int numberNames, ref string[] names) =>
                    sapModel.RespCombo.GetNameList(ref numberNames, ref names),
                (ETABSv1.cSapModel sapModel, string name, int combinationType) =>
                    sapModel.RespCombo.Add(name, combinationType),
                (ETABSv1.cSapModel sapModel, string name) =>
                    sapModel.RespCombo.Delete(name),
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
                (ETABSv1.cSapModel sapModel, string comboName, int caseNameType, string caseName) =>
                    sapModel.RespCombo.DeleteCase(comboName, (ETABSv1.eCNameType)caseNameType, caseName),
                (ETABSv1.cSapModel sapModel, string comboName, int caseNameType, string caseName, double scaleFactor) =>
                {
                    ETABSv1.eCNameType nameType = (ETABSv1.eCNameType)caseNameType;
                    return sapModel.RespCombo.SetCaseList(comboName, ref nameType, caseName, scaleFactor);
                });

            if (result.IsSuccess)
            {
                RefreshView(sapModelResult.Data);
            }

            return result;
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

        public OperationResult<int> GetPresentUnits()
        {
            var sapModelResult = EnsureEtabsSapModel();
            if (!sapModelResult.IsSuccess) return OperationResult<int>.Failure(sapModelResult.Message);
            return OperationResult<int>.Success((int)sapModelResult.Data.GetPresentUnits());
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

        private static OperationResult SelectOutputCasesForTables(
            ETABSv1.cSapModel sapModel,
            IReadOnlyList<CSISapModelOutputCaseDTO> selectedOutputCases)
        {
            int deselectRet = sapModel.Results.Setup.DeselectAllCasesAndCombosForOutput();
            if (deselectRet != 0)
            {
                return OperationResult.Failure($"Failed to clear ETABS output case selection (return code {deselectRet}).");
            }

            foreach (var outputCase in selectedOutputCases)
            {
                if (outputCase == null || string.IsNullOrWhiteSpace(outputCase.Name))
                {
                    continue;
                }

                int selectRet = outputCase.IsLoadCombination
                    ? sapModel.Results.Setup.SetComboSelectedForOutput(outputCase.Name, true)
                    : sapModel.Results.Setup.SetCaseSelectedForOutput(outputCase.Name, true);

                if (selectRet != 0)
                {
                    return OperationResult.Failure($"Failed to select '{outputCase.Name}' for output (return code {selectRet}).");
                }
            }

            string[] selectedCaseNamesForDisplay = CreateSelectedOutputCaseNameArray(selectedOutputCases, false);
            int caseDisplayRet = sapModel.DatabaseTables.SetLoadCasesSelectedForDisplay(ref selectedCaseNamesForDisplay);
            if (caseDisplayRet != 0)
            {
                return OperationResult.Failure($"Failed to select ETABS database table display load cases (return code {caseDisplayRet}).");
            }

            string[] selectedCombinationNamesForDisplay = CreateSelectedOutputCaseNameArray(selectedOutputCases, true);
            int combinationDisplayRet = sapModel.DatabaseTables.SetLoadCombinationsSelectedForDisplay(ref selectedCombinationNamesForDisplay);
            if (combinationDisplayRet != 0)
            {
                return OperationResult.Failure($"Failed to select ETABS database table display load combinations (return code {combinationDisplayRet}).");
            }

            return OperationResult.Success();
        }

        private static OperationResult SelectLoadCasesOnlyForTables(
            ETABSv1.cSapModel sapModel,
            IReadOnlyList<CSISapModelOutputCaseDTO> selectedOutputCases)
        {
            int deselectRet = sapModel.Results.Setup.DeselectAllCasesAndCombosForOutput();
            if (deselectRet != 0)
            {
                return OperationResult.Failure($"Failed to clear ETABS output case selection (return code {deselectRet}).");
            }

            foreach (var outputCase in selectedOutputCases)
            {
                if (outputCase == null || string.IsNullOrWhiteSpace(outputCase.Name))
                {
                    continue;
                }

                if (outputCase.IsLoadCombination)
                {
                    return OperationResult.Failure("Select ETABS load cases only for this table.");
                }

                int selectRet = sapModel.Results.Setup.SetCaseSelectedForOutput(outputCase.Name, true);
                if (selectRet != 0)
                {
                    return OperationResult.Failure($"Failed to select '{outputCase.Name}' for output (return code {selectRet}).");
                }
            }

            string[] selectedCaseNamesForDisplay = CreateSelectedOutputCaseNameArray(selectedOutputCases, false);
            int caseDisplayRet = sapModel.DatabaseTables.SetLoadCasesSelectedForDisplay(ref selectedCaseNamesForDisplay);
            if (caseDisplayRet != 0)
            {
                return OperationResult.Failure($"Failed to select ETABS database table display load cases (return code {caseDisplayRet}).");
            }

            string[] selectedCombinationNamesForDisplay = new string[0];
            int combinationDisplayRet = sapModel.DatabaseTables.SetLoadCombinationsSelectedForDisplay(ref selectedCombinationNamesForDisplay);
            if (combinationDisplayRet != 0)
            {
                return OperationResult.Failure($"Failed to clear ETABS database table display load combinations (return code {combinationDisplayRet}).");
            }

            return OperationResult.Success();
        }

        private delegate int GetCaseNameListDelegate(ref int numberNames, ref string[] names);

        private static void AddModalCasesFromSpecificApi(
            ICollection<CSISapModelOutputCaseDTO> outputCases,
            ISet<string> seenNames,
            string caseType,
            GetCaseNameListDelegate getNameList)
        {
            if (getNameList == null)
            {
                return;
            }

            try
            {
                int numberNames = 0;
                string[] names = null;
                int ret = getNameList(ref numberNames, ref names);
                if (ret == 0)
                {
                    AddModalCaseNames(outputCases, seenNames, names, numberNames, caseType);
                }
            }
            catch
            {
                // Older ETABS interop builds may not expose ModalEigen/ModalRitz GetNameList.
            }
        }

        private static void AddModalCaseNames(
            ICollection<CSISapModelOutputCaseDTO> outputCases,
            ISet<string> seenNames,
            string[] names,
            int numberNames,
            string caseType)
        {
            if (outputCases == null || seenNames == null || names == null)
            {
                return;
            }

            for (int i = 0; i < numberNames && i < names.Length; i++)
            {
                string name = names[i];
                if (string.IsNullOrWhiteSpace(name) || !seenNames.Add(name))
                {
                    continue;
                }

                outputCases.Add(new CSISapModelOutputCaseDTO
                {
                    Name = name,
                    Type = caseType,
                    IsLoadCombination = false
                });
            }
        }

        private static string[] CreateSelectedOutputCaseNameArray(
            IReadOnlyList<CSISapModelOutputCaseDTO> selectedOutputCases,
            bool isLoadCombination)
        {
            var names = new List<string>();
            if (selectedOutputCases == null)
            {
                return names.ToArray();
            }

            foreach (var outputCase in selectedOutputCases)
            {
                if (outputCase != null &&
                    outputCase.IsLoadCombination == isLoadCombination &&
                    !string.IsNullOrWhiteSpace(outputCase.Name))
                {
                    names.Add(outputCase.Name);
                }
            }

            return names.ToArray();
        }

        private static IReadOnlyList<CSISapModelBaseReactionRowDTO> FilterBaseReactionRows(
            IReadOnlyList<CSISapModelBaseReactionRowDTO> rows,
            IReadOnlyList<CSISapModelOutputCaseDTO> selectedOutputCases)
        {
            if (rows == null || rows.Count == 0)
            {
                return rows ?? new List<CSISapModelBaseReactionRowDTO>();
            }

            var selectedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var outputCase in selectedOutputCases ?? new CSISapModelOutputCaseDTO[0])
            {
                if (outputCase != null && !string.IsNullOrWhiteSpace(outputCase.Name))
                {
                    selectedNames.Add(outputCase.Name.Trim());
                }
            }

            if (selectedNames.Count == 0)
            {
                return rows;
            }

            var filteredRows = new List<CSISapModelBaseReactionRowDTO>();
            foreach (var row in rows)
            {
                string outputCaseName = Convert.ToString(row.OutputCase, CultureInfo.InvariantCulture);
                if (string.IsNullOrWhiteSpace(outputCaseName))
                {
                    filteredRows.Add(row);
                    continue;
                }

                if (selectedNames.Contains(outputCaseName.Trim()))
                {
                    filteredRows.Add(row);
                }
            }

            return filteredRows;
        }

        private static IReadOnlyList<CSISapModelBaseReactionRowDTO> ParseBaseReactionRows(string[] fieldsKeysIncluded, int numberRecords, string[] tableData)
        {
            var rows = new List<CSISapModelBaseReactionRowDTO>();
            if (numberRecords <= 0)
            {
                return rows;
            }

            int fieldCount = fieldsKeysIncluded == null ? 0 : fieldsKeysIncluded.Length;
            if (fieldCount == 0)
            {
                return rows;
            }

            var fieldIndexes = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < fieldsKeysIncluded.Length; i++)
            {
                string normalized = NormalizeFieldKey(fieldsKeysIncluded[i]);
                if (!string.IsNullOrWhiteSpace(normalized) && !fieldIndexes.ContainsKey(normalized))
                {
                    fieldIndexes.Add(normalized, i);
                }
            }

            for (int recordIndex = 0; recordIndex < numberRecords; recordIndex++)
            {
                rows.Add(new CSISapModelBaseReactionRowDTO
                {
                    OutputCase = ReadTableValue(fieldIndexes, tableData, fieldCount, recordIndex, false, "Output Case", "OutputCase", "Load Case", "LoadCase"),
                    CaseType = ReadTableValue(fieldIndexes, tableData, fieldCount, recordIndex, false, "Case Type", "CaseType"),
                    StepType = ReadTableValue(fieldIndexes, tableData, fieldCount, recordIndex, false, "Step Type", "StepType"),
                    StepNumber = ReadTableValue(fieldIndexes, tableData, fieldCount, recordIndex, true, "Step Number", "Step Num", "StepNum", "Step"),
                    FX = ReadTableValue(fieldIndexes, tableData, fieldCount, recordIndex, true, "FX", "F1"),
                    FY = ReadTableValue(fieldIndexes, tableData, fieldCount, recordIndex, true, "FY", "F2"),
                    FZ = ReadTableValue(fieldIndexes, tableData, fieldCount, recordIndex, true, "FZ", "F3"),
                    MX = ReadTableValue(fieldIndexes, tableData, fieldCount, recordIndex, true, "MX", "M1"),
                    MY = ReadTableValue(fieldIndexes, tableData, fieldCount, recordIndex, true, "MY", "M2"),
                    MZ = ReadTableValue(fieldIndexes, tableData, fieldCount, recordIndex, true, "MZ", "M3"),
                    X = ReadTableValue(fieldIndexes, tableData, fieldCount, recordIndex, true, "X"),
                    Y = ReadTableValue(fieldIndexes, tableData, fieldCount, recordIndex, true, "Y"),
                    Z = ReadTableValue(fieldIndexes, tableData, fieldCount, recordIndex, true, "Z")
                });
            }

            return rows;
        }

        private static IReadOnlyList<CSISapModelModalMassParticipationRowDTO> FilterModalMassParticipationRows(
            IReadOnlyList<CSISapModelModalMassParticipationRowDTO> rows,
            IReadOnlyList<CSISapModelOutputCaseDTO> selectedOutputCases)
        {
            if (rows == null || rows.Count == 0)
            {
                return rows ?? new List<CSISapModelModalMassParticipationRowDTO>();
            }

            var selectedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var outputCase in selectedOutputCases ?? new CSISapModelOutputCaseDTO[0])
            {
                if (outputCase != null && !string.IsNullOrWhiteSpace(outputCase.Name))
                {
                    selectedNames.Add(outputCase.Name.Trim());
                }
            }

            if (selectedNames.Count == 0)
            {
                return rows;
            }

            var filteredRows = new List<CSISapModelModalMassParticipationRowDTO>();
            foreach (var row in rows)
            {
                string outputCaseName = Convert.ToString(row.OutputCase, CultureInfo.InvariantCulture);
                if (string.IsNullOrWhiteSpace(outputCaseName) || selectedNames.Contains(outputCaseName.Trim()))
                {
                    filteredRows.Add(row);
                }
            }

            return filteredRows;
        }

        private static IReadOnlyList<CSISapModelModalMassParticipationRowDTO> ParseModalMassParticipationRows(string[] fieldsKeysIncluded, int numberRecords, string[] tableData)
        {
            var rows = new List<CSISapModelModalMassParticipationRowDTO>();
            if (numberRecords <= 0)
            {
                return rows;
            }

            int fieldCount = fieldsKeysIncluded == null ? 0 : fieldsKeysIncluded.Length;
            if (fieldCount == 0)
            {
                return rows;
            }

            var fieldIndexes = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < fieldsKeysIncluded.Length; i++)
            {
                string normalized = NormalizeFieldKey(fieldsKeysIncluded[i]);
                if (!string.IsNullOrWhiteSpace(normalized) && !fieldIndexes.ContainsKey(normalized))
                {
                    fieldIndexes.Add(normalized, i);
                }
            }

            for (int recordIndex = 0; recordIndex < numberRecords; recordIndex++)
            {
                rows.Add(new CSISapModelModalMassParticipationRowDTO
                {
                    OutputCase = ReadTableValue(fieldIndexes, tableData, fieldCount, recordIndex, false, "Output Case", "OutputCase", "Load Case", "LoadCase", "Case"),
                    StepType = ReadTableValue(fieldIndexes, tableData, fieldCount, recordIndex, false, "Step Type", "StepType"),
                    StepNumber = ReadTableValue(fieldIndexes, tableData, fieldCount, recordIndex, true, "Step Number", "Step Num", "StepNum", "Mode"),
                    Period = ReadTableValue(fieldIndexes, tableData, fieldCount, recordIndex, true, "Period"),
                    UX = ReadTableValue(fieldIndexes, tableData, fieldCount, recordIndex, true, "UX", "U1"),
                    UY = ReadTableValue(fieldIndexes, tableData, fieldCount, recordIndex, true, "UY", "U2"),
                    UZ = ReadTableValue(fieldIndexes, tableData, fieldCount, recordIndex, true, "UZ", "U3"),
                    SumUX = ReadTableValue(fieldIndexes, tableData, fieldCount, recordIndex, true, "Sum UX", "SumUX", "Sum U1", "SumU1"),
                    SumUY = ReadTableValue(fieldIndexes, tableData, fieldCount, recordIndex, true, "Sum UY", "SumUY", "Sum U2", "SumU2"),
                    SumUZ = ReadTableValue(fieldIndexes, tableData, fieldCount, recordIndex, true, "Sum UZ", "SumUZ", "Sum U3", "SumU3"),
                    RX = ReadTableValue(fieldIndexes, tableData, fieldCount, recordIndex, true, "RX", "R1"),
                    RY = ReadTableValue(fieldIndexes, tableData, fieldCount, recordIndex, true, "RY", "R2"),
                    RZ = ReadTableValue(fieldIndexes, tableData, fieldCount, recordIndex, true, "RZ", "R3"),
                    SumRX = ReadTableValue(fieldIndexes, tableData, fieldCount, recordIndex, true, "Sum RX", "SumRX", "Sum R1", "SumR1"),
                    SumRY = ReadTableValue(fieldIndexes, tableData, fieldCount, recordIndex, true, "Sum RY", "SumRY", "Sum R2", "SumR2"),
                    SumRZ = ReadTableValue(fieldIndexes, tableData, fieldCount, recordIndex, true, "Sum RZ", "SumRZ", "Sum R3", "SumR3")
                });
            }

            return rows;
        }

        private static IReadOnlyList<CSISapModelStoryForceRowDTO> FilterStoryForceRows(
            IReadOnlyList<CSISapModelStoryForceRowDTO> rows,
            IReadOnlyList<CSISapModelOutputCaseDTO> selectedOutputCases)
        {
            if (rows == null || rows.Count == 0)
            {
                return rows ?? new List<CSISapModelStoryForceRowDTO>();
            }

            var selectedNames = CreateSelectedOutputCaseSet(selectedOutputCases);
            if (selectedNames.Count == 0)
            {
                return rows;
            }

            var filteredRows = new List<CSISapModelStoryForceRowDTO>();
            foreach (var row in rows)
            {
                string outputCaseName = Convert.ToString(row.OutputCase, CultureInfo.InvariantCulture);
                if (string.IsNullOrWhiteSpace(outputCaseName) || selectedNames.Contains(outputCaseName.Trim()))
                {
                    filteredRows.Add(row);
                }
            }

            return filteredRows;
        }

        private static IReadOnlyList<CSISapModelStoryDisplacementRowDTO> FilterStoryDisplacementRows(
            IReadOnlyList<CSISapModelStoryDisplacementRowDTO> rows,
            IReadOnlyList<CSISapModelOutputCaseDTO> selectedOutputCases)
        {
            if (rows == null || rows.Count == 0)
            {
                return rows ?? new List<CSISapModelStoryDisplacementRowDTO>();
            }

            var selectedNames = CreateSelectedOutputCaseSet(selectedOutputCases);
            if (selectedNames.Count == 0)
            {
                return rows;
            }

            var filteredRows = new List<CSISapModelStoryDisplacementRowDTO>();
            foreach (var row in rows)
            {
                string outputCaseName = Convert.ToString(row.OutputCase, CultureInfo.InvariantCulture);
                if (string.IsNullOrWhiteSpace(outputCaseName) || selectedNames.Contains(outputCaseName.Trim()))
                {
                    filteredRows.Add(row);
                }
            }

            return filteredRows;
        }

        private static HashSet<string> CreateSelectedOutputCaseSet(IReadOnlyList<CSISapModelOutputCaseDTO> selectedOutputCases)
        {
            var selectedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var outputCase in selectedOutputCases ?? new CSISapModelOutputCaseDTO[0])
            {
                if (outputCase != null && !string.IsNullOrWhiteSpace(outputCase.Name))
                {
                    selectedNames.Add(outputCase.Name.Trim());
                }
            }

            return selectedNames;
        }

        private static CSISapModelDisplayTableDTO ParseDisplayTable(string[] fields, int numberRecords, string[] tableData)
        {
            var fieldKeys = new List<string>(fields ?? new string[0]);
            var rows = new List<object[]>();
            if (fieldKeys.Count == 0 || numberRecords <= 0 || tableData == null)
            {
                return new CSISapModelDisplayTableDTO { FieldKeys = fieldKeys, Rows = rows };
            }

            for (int recordIndex = 0; recordIndex < numberRecords; recordIndex++)
            {
                var row = new object[fieldKeys.Count];
                for (int fieldIndex = 0; fieldIndex < fieldKeys.Count; fieldIndex++)
                {
                    int dataIndex = recordIndex * fieldKeys.Count + fieldIndex;
                    string value = dataIndex >= 0 && dataIndex < tableData.Length ? tableData[dataIndex] : string.Empty;
                    row[fieldIndex] = ParseDisplayTableValue(fieldKeys[fieldIndex], value);
                }

                rows.Add(row);
            }

            return new CSISapModelDisplayTableDTO { FieldKeys = fieldKeys, Rows = rows };
        }

        private static CSISapModelDisplayTableDTO FilterDisplayTableRows(
            CSISapModelDisplayTableDTO table,
            IReadOnlyList<CSISapModelOutputCaseDTO> selectedOutputCases)
        {
            return FilterDisplayTableRows(table, selectedOutputCases, null);
        }

        private static CSISapModelDisplayTableDTO FilterDisplayTableRows(
            CSISapModelDisplayTableDTO table,
            IReadOnlyList<CSISapModelOutputCaseDTO> selectedOutputCases,
            string displayTableName)
        {
            if (table == null || table.Rows == null || table.Rows.Count == 0)
            {
                return table ?? new CSISapModelDisplayTableDTO
                {
                    FieldKeys = new List<string>(),
                    Rows = new List<object[]>()
                };
            }

            int outputCaseIndex = IsResponseSpectrumModalInfoTable(displayTableName)
                ? FindFieldIndex(
                    table.FieldKeys,
                    "Response Spectrum Case",
                    "ResponseSpectrumCase",
                    "Spectrum Case",
                    "SpectrumCase",
                    "Spectrum Load Case",
                    "SpectrumLoadCase",
                    "Response Spectrum Load Case",
                    "ResponseSpectrumLoadCase",
                    "Spec Case",
                    "SpecCase",
                    "RS Case",
                    "RSCase",
                    "Output Case",
                    "OutputCase",
                    "Output Case Name",
                    "OutputCaseName",
                    "Load Case",
                    "LoadCase",
                    "Load Case Name",
                    "LoadCaseName",
                    "Case",
                    "Case Name",
                    "CaseName")
                : FindFieldIndex(
                    table.FieldKeys,
                    "Output Case",
                    "OutputCase",
                    "Output Case Name",
                    "OutputCaseName",
                    "Load Case",
                    "LoadCase",
                    "Load Case Name",
                    "LoadCaseName",
                    "Modal Case",
                    "ModalCase",
                    "Case",
                    "Case Name",
                    "CaseName");
            if (outputCaseIndex < 0)
            {
                return table;
            }

            var selectedNames = CreateSelectedOutputCaseSet(selectedOutputCases);
            if (selectedNames.Count == 0)
            {
                return table;
            }

            var filteredRows = new List<object[]>();
            foreach (object[] row in table.Rows)
            {
                string outputCaseName = row != null && outputCaseIndex < row.Length
                    ? Convert.ToString(row[outputCaseIndex], CultureInfo.InvariantCulture)
                    : string.Empty;
                if (!string.IsNullOrWhiteSpace(outputCaseName) && selectedNames.Contains(outputCaseName.Trim()))
                {
                    filteredRows.Add(row);
                }
            }

            return new CSISapModelDisplayTableDTO { FieldKeys = table.FieldKeys, Rows = filteredRows };
        }

        private static int FindFieldIndex(IReadOnlyList<string> fields, params string[] aliases)
        {
            if (fields == null || aliases == null)
            {
                return -1;
            }

            for (int fieldIndex = 0; fieldIndex < fields.Count; fieldIndex++)
            {
                foreach (string alias in aliases)
                {
                    if (string.Equals(NormalizeFieldKey(fields[fieldIndex]), NormalizeFieldKey(alias), StringComparison.OrdinalIgnoreCase))
                    {
                        return fieldIndex;
                    }
                }
            }

            return -1;
        }

        private static object ParseDisplayTableValue(string fieldKey, string value)
        {
            if (IsDisplayTableTextField(fieldKey))
            {
                return value ?? string.Empty;
            }

            double number;
            if (double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out number) ||
                double.TryParse(value, NumberStyles.Float, CultureInfo.CurrentCulture, out number))
            {
                return number;
            }

            return value ?? string.Empty;
        }

        private static bool IsDisplayTableTextField(string fieldKey)
        {
            string normalizedFieldKey = NormalizeFieldKey(fieldKey);
            return normalizedFieldKey == NormalizeFieldKey("Story") ||
                   normalizedFieldKey == NormalizeFieldKey("Story Name") ||
                   normalizedFieldKey == NormalizeFieldKey("Output Case") ||
                   normalizedFieldKey == NormalizeFieldKey("Output Case Name") ||
                   normalizedFieldKey == NormalizeFieldKey("Load Case") ||
                   normalizedFieldKey == NormalizeFieldKey("Load Case Name") ||
                   normalizedFieldKey == NormalizeFieldKey("Modal Case") ||
                   normalizedFieldKey == NormalizeFieldKey("Response Spectrum Case") ||
                   normalizedFieldKey == NormalizeFieldKey("Spectrum Case") ||
                   normalizedFieldKey == NormalizeFieldKey("Spectrum Load Case") ||
                   normalizedFieldKey == NormalizeFieldKey("Response Spectrum Load Case") ||
                   normalizedFieldKey == NormalizeFieldKey("Spec Case") ||
                   normalizedFieldKey == NormalizeFieldKey("RS Case") ||
                   normalizedFieldKey == NormalizeFieldKey("Case") ||
                   normalizedFieldKey == NormalizeFieldKey("Case Name") ||
                   normalizedFieldKey == NormalizeFieldKey("Case Type") ||
                   normalizedFieldKey == NormalizeFieldKey("Step Type") ||
                   normalizedFieldKey == NormalizeFieldKey("Direction") ||
                   normalizedFieldKey == NormalizeFieldKey("Location") ||
                   normalizedFieldKey == NormalizeFieldKey("Label");
        }

        private static IReadOnlyList<CSISapModelStoryForceRowDTO> ParseStoryForceRows(string[] fieldsKeysIncluded, int numberRecords, string[] tableData)
        {
            var rows = new List<CSISapModelStoryForceRowDTO>();
            if (numberRecords <= 0)
            {
                return rows;
            }

            int fieldCount = fieldsKeysIncluded == null ? 0 : fieldsKeysIncluded.Length;
            if (fieldCount == 0)
            {
                return rows;
            }

            var fieldIndexes = CreateFieldIndexMap(fieldsKeysIncluded);
            for (int recordIndex = 0; recordIndex < numberRecords; recordIndex++)
            {
                rows.Add(new CSISapModelStoryForceRowDTO
                {
                    Story = ReadTableValue(fieldIndexes, tableData, fieldCount, recordIndex, false, "Story", "Story Name", "StoryName"),
                    OutputCase = ReadTableValue(fieldIndexes, tableData, fieldCount, recordIndex, false, "Output Case", "OutputCase", "Load Case", "LoadCase", "Case"),
                    CaseType = ReadTableValue(fieldIndexes, tableData, fieldCount, recordIndex, false, "Case Type", "CaseType"),
                    StepType = ReadTableValue(fieldIndexes, tableData, fieldCount, recordIndex, false, "Step Type", "StepType"),
                    StepNumber = ReadTableValue(fieldIndexes, tableData, fieldCount, recordIndex, true, "Step Number", "Step Num", "StepNum", "Step"),
                    Location = ReadTableValue(fieldIndexes, tableData, fieldCount, recordIndex, false, "Location"),
                    P = ReadTableValue(fieldIndexes, tableData, fieldCount, recordIndex, true, "P", "FZ", "Axial"),
                    VX = ReadTableValue(fieldIndexes, tableData, fieldCount, recordIndex, true, "VX", "Vx", "FX", "Shear X", "ShearX"),
                    VY = ReadTableValue(fieldIndexes, tableData, fieldCount, recordIndex, true, "VY", "Vy", "FY", "Shear Y", "ShearY"),
                    T = ReadTableValue(fieldIndexes, tableData, fieldCount, recordIndex, true, "T", "Torsion", "MZ"),
                    MX = ReadTableValue(fieldIndexes, tableData, fieldCount, recordIndex, true, "MX", "Mx"),
                    MY = ReadTableValue(fieldIndexes, tableData, fieldCount, recordIndex, true, "MY", "My")
                });
            }

            return rows;
        }

        private static IReadOnlyList<CSISapModelStoryDisplacementRowDTO> ParseStoryDisplacementRows(string[] fieldsKeysIncluded, int numberRecords, string[] tableData)
        {
            var rows = new List<CSISapModelStoryDisplacementRowDTO>();
            if (numberRecords <= 0)
            {
                return rows;
            }

            int fieldCount = fieldsKeysIncluded == null ? 0 : fieldsKeysIncluded.Length;
            if (fieldCount == 0)
            {
                return rows;
            }

            var fieldIndexes = CreateFieldIndexMap(fieldsKeysIncluded);
            for (int recordIndex = 0; recordIndex < numberRecords; recordIndex++)
            {
                rows.Add(new CSISapModelStoryDisplacementRowDTO
                {
                    Story = ReadTableValue(fieldIndexes, tableData, fieldCount, recordIndex, false, "Story", "Story Name", "StoryName"),
                    OutputCase = ReadTableValue(fieldIndexes, tableData, fieldCount, recordIndex, false, "Output Case", "OutputCase", "Load Case", "LoadCase", "Case"),
                    CaseType = ReadTableValue(fieldIndexes, tableData, fieldCount, recordIndex, false, "Case Type", "CaseType"),
                    StepType = ReadTableValue(fieldIndexes, tableData, fieldCount, recordIndex, false, "Step Type", "StepType"),
                    StepNumber = ReadTableValue(fieldIndexes, tableData, fieldCount, recordIndex, true, "Step Number", "Step Num", "StepNum", "Step"),
                    UX = ReadTableValue(fieldIndexes, tableData, fieldCount, recordIndex, true, "UX", "U1", "X"),
                    UY = ReadTableValue(fieldIndexes, tableData, fieldCount, recordIndex, true, "UY", "U2", "Y"),
                    UZ = ReadTableValue(fieldIndexes, tableData, fieldCount, recordIndex, true, "UZ", "U3", "Z"),
                    RX = ReadTableValue(fieldIndexes, tableData, fieldCount, recordIndex, true, "RX", "R1"),
                    RY = ReadTableValue(fieldIndexes, tableData, fieldCount, recordIndex, true, "RY", "R2"),
                    RZ = ReadTableValue(fieldIndexes, tableData, fieldCount, recordIndex, true, "RZ", "R3")
                });
            }

            return rows;
        }

        private static Dictionary<string, int> CreateFieldIndexMap(string[] fieldsKeysIncluded)
        {
            var fieldIndexes = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            if (fieldsKeysIncluded == null)
            {
                return fieldIndexes;
            }

            for (int i = 0; i < fieldsKeysIncluded.Length; i++)
            {
                string normalized = NormalizeFieldKey(fieldsKeysIncluded[i]);
                if (!string.IsNullOrWhiteSpace(normalized) && !fieldIndexes.ContainsKey(normalized))
                {
                    fieldIndexes.Add(normalized, i);
                }
            }

            return fieldIndexes;
        }

        private static object ReadTableValue(
            IDictionary<string, int> fieldIndexes,
            string[] tableData,
            int fieldCount,
            int recordIndex,
            bool numeric,
            params string[] aliases)
        {
            if (fieldIndexes == null || aliases == null || tableData == null)
            {
                return string.Empty;
            }

            foreach (string alias in aliases)
            {
                int fieldIndex;
                if (!fieldIndexes.TryGetValue(NormalizeFieldKey(alias), out fieldIndex))
                {
                    continue;
                }

                int dataIndex = recordIndex * fieldCount + fieldIndex;
                if (dataIndex < 0 || dataIndex >= tableData.Length)
                {
                    return string.Empty;
                }

                string value = tableData[dataIndex];
                if (!numeric)
                {
                    return value ?? string.Empty;
                }

                double number;
                if (double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out number) ||
                    double.TryParse(value, NumberStyles.Float, CultureInfo.CurrentCulture, out number))
                {
                    return number;
                }

                return value ?? string.Empty;
            }

            return string.Empty;
        }

        private static string NormalizeFieldKey(string fieldKey)
        {
            if (string.IsNullOrWhiteSpace(fieldKey))
            {
                return string.Empty;
            }

            char[] chars = fieldKey.ToCharArray();
            var normalized = new System.Text.StringBuilder(chars.Length);
            foreach (char c in chars)
            {
                if (char.IsLetterOrDigit(c))
                {
                    normalized.Append(char.ToUpperInvariant(c));
                }
            }

            return normalized.ToString();
        }

        private static string FormatResponseCombinationType(int type)
        {
            switch (type)
            {
                case 0: return "Linear Add";
                case 1: return "Envelope";
                case 2: return "Absolute Add";
                case 3: return "SRSS";
                case 4: return "Range Add";
                default: return type.ToString(CultureInfo.InvariantCulture);
            }
        }

        private static string FormatLoadCaseType(ETABSv1.eLoadCaseType caseType)
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

        private static bool[] ToReleaseArray(IReadOnlyList<bool> releases)
        {
            var values = new bool[6];
            if (releases == null)
            {
                return values;
            }

            for (int i = 0; i < releases.Count && i < values.Length; i++)
            {
                values[i] = releases[i];
            }

            return values;
        }

    }
}
