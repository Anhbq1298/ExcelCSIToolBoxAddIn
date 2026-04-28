using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using ExcelCSIToolBoxAddIn.Adapters;
using ExcelCSIToolBoxAddIn.Common.Results;
using ExcelCSIToolBoxAddIn.Core.Geometry;
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
        private readonly ICsiModelAdapter _modelAdapter;
        private CSISapModelConnectionInfo _currentConnection;

        public Sap2000ConnectionService()
            : this(new Sap2000ModelAdapter())
        {
        }

        public Sap2000ConnectionService(ICsiModelAdapter modelAdapter)
        {
            _modelAdapter = modelAdapter ?? throw new ArgumentNullException(nameof(modelAdapter));
        }

        public string ProductName => "SAP2000";

        public OperationResult<CSISapModelConnectionInfo> TryAttachToRunningInstance()
        {
            var attachResult = _modelAdapter.AttachToRunningInstance();
            if (!attachResult.IsSuccess)
            {
                _currentConnection = null;
                return OperationResult<CSISapModelConnectionInfo>.Failure(attachResult.Message);
            }

            var sapObject = attachResult.ApplicationObject as SAP2000v1.cOAPI;
            var sapModel = attachResult.SapModel as SAP2000v1.cSapModel;
            if (sapObject == null || sapModel == null)
            {
                _currentConnection = null;
                return OperationResult<CSISapModelConnectionInfo>.Failure("The attached SAP2000 instance is invalid. Please reattach and try again.");
            }

            try
            {
                string modelPath = string.Empty;
                string modelName = "Unsaved Model";
                try
                {
                    modelPath = sapModel.GetModelFilename(true);
                    modelName = string.IsNullOrWhiteSpace(modelPath)
                        ? "Unsaved Model"
                        : Path.GetFileName(modelPath);
                }
                catch
                {
                }

                string modelCurrentUnit = "Units unavailable";
                try
                {
                    modelCurrentUnit = Sap2000UnitFormatter.FormatPresentUnits(sapModel.GetPresentUnits());
                }
                catch
                {
                }

                _currentConnection = new CSISapModelConnectionInfo
                {
                    IsConnected = true,
                    ModelPath = modelPath,
                    ModelFileName = modelName,
                    ModelCurrentUnit = modelCurrentUnit,
                    CsiObject = sapObject,
                    SapModel = sapModel
                };

                return OperationResult<CSISapModelConnectionInfo>.Success(_currentConnection);
            }
            catch
            {
                _currentConnection = null;
                return OperationResult<CSISapModelConnectionInfo>.Failure("Failed to attach to the running SAP2000 instance.");
            }
        }

        public OperationResult<CSISapModelConnectionInfo> GetCurrentConnection()
        {
            if (_currentConnection?.SapModel == null)
            {
                return OperationResult<CSISapModelConnectionInfo>.Failure("No SAP2000 model is currently connected. Please click Attach.");
            }

            return OperationResult<CSISapModelConnectionInfo>.Success(_currentConnection);
        }

        public OperationResult CloseCurrentInstance()
        {
            if (_currentConnection?.CsiObject == null)
            {
                return OperationResult.Failure("No running SAP2000 instance is currently attached.");
            }

            try
            {
                var sapApplication = _currentConnection.CsiObject as SAP2000v1.cOAPI;
                if (sapApplication == null)
                {
                    return OperationResult.Failure("The attached SAP2000 instance is invalid. Please reattach and try again.");
                }

                int result = sapApplication.ApplicationExit(false);
                if (result != 0)
                {
                    return OperationResult.Failure($"SAP2000 failed to close the attached instance (ApplicationExit returned {result}).");
                }

                ResetCurrentConnection();
                return OperationResult.Success("Successfully closed the attached SAP2000 instance.");
            }
            catch (COMException ex)
            {
                return OperationResult.Failure($"SAP2000 COM error while closing attached instance: {ex.Message}");
            }
            catch
            {
                return OperationResult.Failure("Failed to close the attached SAP2000 instance.");
            }
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

        public OperationResult<CSISapModelAddPointsResult> AddPointsByCartesian(IReadOnlyList<CSISapModelPointCartesianInput> pointInputs)
        {
            var sapModelResult = EnsureSap2000SapModel();
            if (!sapModelResult.IsSuccess)
            {
                return OperationResult<CSISapModelAddPointsResult>.Failure(sapModelResult.Message);
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

        public OperationResult<CSISapModelAddFramesResult> AddFramesByCoordinates(IReadOnlyList<CSISapModelFrameByCoordInput> frameInputs)
        {
            var sapModelResult = EnsureSap2000SapModel();
            if (!sapModelResult.IsSuccess)
            {
                return OperationResult<CSISapModelAddFramesResult>.Failure(sapModelResult.Message);
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

        public OperationResult<CSISapModelAddFramesResult> AddFramesByPoint(IReadOnlyList<CSISapModelFrameByPointInput> frameInputs)
        {
            var sapModelResult = EnsureSap2000SapModel();
            if (!sapModelResult.IsSuccess)
            {
                return OperationResult<CSISapModelAddFramesResult>.Failure(sapModelResult.Message);
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

        public OperationResult<IReadOnlyList<CSISapModelPointData>> GetSelectedPointsFromActiveModel()
        {
            var sapModelResult = EnsureSap2000SapModel();
            if (!sapModelResult.IsSuccess)
            {
                return OperationResult<IReadOnlyList<CSISapModelPointData>>.Failure(sapModelResult.Message);
            }

            return CSISapModelPointObjectService.GetSelectedPointsFromActiveModel(
                ProductName,
                sapModelResult.Data,
                (SAP2000v1.cSapModel sapModel, ref int numberItems, ref int[] objectTypes, ref string[] objectNames) =>
                    sapModel.SelectObj.GetSelected(ref numberItems, ref objectTypes, ref objectNames),
                (SAP2000v1.cSapModel sapModel, string pointName, ref double x, ref double y, ref double z) =>
                    sapModel.PointObj.GetCoordCartesian(pointName, ref x, ref y, ref z, "Global"),
                null);
        }

        public OperationResult<IReadOnlyList<string>> GetSelectedFramesFromActiveModel()
        {
            var sapModelResult = EnsureSap2000SapModel();
            if (!sapModelResult.IsSuccess)
            {
                return OperationResult<IReadOnlyList<string>>.Failure(sapModelResult.Message);
            }

            return CSISapModelFrameObjectService.GetSelectedFramesFromActiveModel(
                ProductName,
                sapModelResult.Data,
                (SAP2000v1.cSapModel sapModel, ref int numberItems, ref int[] objectTypes, ref string[] objectNames) =>
                    sapModel.SelectObj.GetSelected(ref numberItems, ref objectTypes, ref objectNames));
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

            return CSISapModelShellObjectService.CreateShellAreasFromSelectedFrames(
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
        }

        private OperationResult<CSISapModelConnectionInfo> EnsureConnection()
        {
            var connectionResult = GetCurrentConnection();
            if (!connectionResult.IsSuccess || connectionResult.Data?.SapModel == null)
            {
                connectionResult = TryAttachToRunningInstance();
            }

            return connectionResult;
        }

        private OperationResult<SAP2000v1.cSapModel> EnsureSap2000SapModel()
        {
            var connectionResult = EnsureConnection();
            if (!connectionResult.IsSuccess || connectionResult.Data?.SapModel == null)
            {
                return OperationResult<SAP2000v1.cSapModel>.Failure(connectionResult.Message);
            }

            var sapModel = connectionResult.Data.SapModel as SAP2000v1.cSapModel;
            if (sapModel == null)
            {
                return OperationResult<SAP2000v1.cSapModel>.Failure("The attached SAP2000 SapModel is invalid. Please reattach and try again.");
            }

            return OperationResult<SAP2000v1.cSapModel>.Success(sapModel);
        }

        private void ResetCurrentConnection()
        {
            if (_currentConnection == null)
            {
                return;
            }

            ReleaseComReference(_currentConnection.SapModel);
            ReleaseComReference(_currentConnection.CsiObject);

            _currentConnection.SapModel = null;
            _currentConnection.CsiObject = null;
            _currentConnection = null;
        }

        private static void ReleaseComReference(object comReference)
        {
            if (comReference == null || !Marshal.IsComObject(comReference))
            {
                return;
            }

            try
            {
                Marshal.FinalReleaseComObject(comReference);
            }
            catch
            {
            }
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


