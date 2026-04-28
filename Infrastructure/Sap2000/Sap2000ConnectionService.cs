using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using ExcelCSIToolBoxAddIn.Adapters;
using ExcelCSIToolBoxAddIn.Common.Results;
using ExcelCSIToolBoxAddIn.Core.Geometry;
using ExcelCSIToolBoxAddIn.Infrastructure.CSISapModel;
using ExcelCSIToolBoxAddIn.UI.Views;
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

            return CSISapModelOperationRunner.SelectObjectsByUniqueNames(
                uniqueNames,
                "point",
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

            return CSISapModelOperationRunner.SelectObjectsByUniqueNames(
                uniqueNames,
                "frame",
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

            return CSISapModelOperationRunner.AddPointsByCartesian(
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

            return CSISapModelOperationRunner.AddFramesByCoordinates(
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

            return CSISapModelOperationRunner.AddFramesByPoint(
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
            var connectionResult = EnsureConnection();
            if (!connectionResult.IsSuccess || connectionResult.Data?.SapModel == null)
            {
                return OperationResult<IReadOnlyList<CSISapModelPointData>>.Failure(connectionResult.Message);
            }

            try
            {
                SAP2000v1.cSapModel sapModel = (SAP2000v1.cSapModel)connectionResult.Data.SapModel;
                int numberItems = 0;
                int[] objectTypes = null;
                string[] objectNames = null;
                int selectedResult = sapModel.SelectObj.GetSelected(ref numberItems, ref objectTypes, ref objectNames);
                if (selectedResult != 0)
                {
                    return OperationResult<IReadOnlyList<CSISapModelPointData>>.Failure("Failed to read selected objects from SAP2000.");
                }

                var points = new List<CSISapModelPointData>();
                for (int i = 0; i < numberItems; i++)
                {
                    if (objectTypes == null || objectNames == null || i >= objectTypes.Length || i >= objectNames.Length)
                    {
                        continue;
                    }

                    if (objectTypes[i] != CSISapModelObjectTypeIds.Point || string.IsNullOrWhiteSpace(objectNames[i]))
                    {
                        continue;
                    }

                    double x = 0;
                    double y = 0;
                    double z = 0;
                    int pointResult = sapModel.PointObj.GetCoordCartesian(objectNames[i], ref x, ref y, ref z, "Global");
                    if (pointResult == 0)
                    {
                        points.Add(new CSISapModelPointData
                        {
                            PointUniqueName = objectNames[i],
                            PointLabel = objectNames[i],
                            X = x,
                            Y = y,
                            Z = z
                        });
                    }
                }

                if (points.Count == 0)
                {
                    return OperationResult<IReadOnlyList<CSISapModelPointData>>.Failure("No point objects are selected in SAP2000.");
                }

                return OperationResult<IReadOnlyList<CSISapModelPointData>>.Success(points);
            }
            catch
            {
                return OperationResult<IReadOnlyList<CSISapModelPointData>>.Failure("Unable to read selected SAP2000 points.");
            }
        }

        public OperationResult<IReadOnlyList<string>> GetSelectedFramesFromActiveModel()
        {
            var connectionResult = EnsureConnection();
            if (!connectionResult.IsSuccess || connectionResult.Data?.SapModel == null)
            {
                return OperationResult<IReadOnlyList<string>>.Failure(connectionResult.Message);
            }

            try
            {
                SAP2000v1.cSapModel sapModel = (SAP2000v1.cSapModel)connectionResult.Data.SapModel;
                int numberItems = 0;
                int[] objectTypes = null;
                string[] objectNames = null;
                int selectedResult = sapModel.SelectObj.GetSelected(ref numberItems, ref objectTypes, ref objectNames);
                if (selectedResult != 0)
                {
                    return OperationResult<IReadOnlyList<string>>.Failure("Failed to read selected objects from SAP2000.");
                }

                var frameUniqueNames = new List<string>();
                for (int i = 0; i < numberItems; i++)
                {
                    if (objectTypes == null || objectNames == null || i >= objectTypes.Length || i >= objectNames.Length)
                    {
                        continue;
                    }

                    var frameUniqueName = objectNames[i];
                    if (objectTypes[i] == CSISapModelObjectTypeIds.Frame && !string.IsNullOrWhiteSpace(frameUniqueName))
                    {
                        frameUniqueNames.Add(frameUniqueName);
                    }
                }

                if (frameUniqueNames.Count == 0)
                {
                    return OperationResult<IReadOnlyList<string>>.Failure("No frame objects are selected in SAP2000.");
                }

                return OperationResult<IReadOnlyList<string>>.Success(frameUniqueNames);
            }
            catch
            {
                return OperationResult<IReadOnlyList<string>>.Failure("Unable to read selected SAP2000 frames.");
            }
        }

        public OperationResult AddSteelISections(IReadOnlyList<CSISapModelSteelISectionInput> inputs)
        {
            return CreateSections(inputs, "Creating Steel I-Sections...", (sapModel, input) =>
                sapModel.PropFrame.SetISection(input.SectionName, input.MaterialName, input.H, input.B, input.Tf, input.Tw, input.B, input.Tf, -1, "", ""));
        }

        public OperationResult AddSteelChannelSections(IReadOnlyList<CSISapModelSteelChannelSectionInput> inputs)
        {
            return CreateSections(inputs, "Creating Steel Channel Sections...", (sapModel, input) =>
                sapModel.PropFrame.SetChannel(input.SectionName, input.MaterialName, input.H, input.B, input.Tf, input.Tw, -1, "", ""));
        }

        public OperationResult AddSteelAngleSections(IReadOnlyList<CSISapModelSteelAngleSectionInput> inputs)
        {
            return CreateSections(inputs, "Creating Steel Angle Sections...", (sapModel, input) =>
                sapModel.PropFrame.SetAngle(input.SectionName, input.MaterialName, input.H, input.B, input.Tf, input.Tw, -1, "", ""));
        }

        public OperationResult AddSteelPipeSections(IReadOnlyList<CSISapModelSteelPipeSectionInput> inputs)
        {
            return CreateSections(inputs, "Creating Steel Pipe Sections...", (sapModel, input) =>
                sapModel.PropFrame.SetPipe(input.SectionName, input.MaterialName, input.OutsideDiameter, input.WallThickness, -1, "", ""));
        }

        public OperationResult AddSteelTubeSections(IReadOnlyList<CSISapModelSteelTubeSectionInput> inputs)
        {
            return CreateSections(inputs, "Creating Steel Tube Sections...", (sapModel, input) =>
                sapModel.PropFrame.SetTube_1(input.SectionName, input.MaterialName, input.H, input.B, input.T, input.T, 0.000000001, -1, "", ""));
        }

        public OperationResult AddConcreteRectangleSections(IReadOnlyList<CSISapModelConcreteRectangleSectionInput> inputs)
        {
            return CreateSections(inputs, "Creating Concrete Rectangle Sections...", (sapModel, input) =>
                sapModel.PropFrame.SetRectangle(input.SectionName, input.MaterialName, input.H, input.B, -1, "", ""));
        }

        public OperationResult AddConcreteCircleSections(IReadOnlyList<CSISapModelConcreteCircleSectionInput> inputs)
        {
            return CreateSections(inputs, "Creating Concrete Circle Sections...", (sapModel, input) =>
                sapModel.PropFrame.SetCircle(input.SectionName, input.MaterialName, input.D, -1, "", ""));
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

            return CSISapModelShellAreaService.CreateShellAreasFromSelectedFrames(
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

        private OperationResult CreateSections<T>(
            IReadOnlyList<T> inputs,
            string title,
            Func<SAP2000v1.cSapModel, T, int> createSection)
        {
            if (inputs == null || inputs.Count == 0)
            {
                return OperationResult.Failure("No valid rows were found in the selected range.");
            }

            var connectionResult = EnsureConnection();
            if (!connectionResult.IsSuccess || connectionResult.Data?.SapModel == null)
            {
                return OperationResult.Failure(connectionResult.Message);
            }

            SAP2000v1.cSapModel sapModel = (SAP2000v1.cSapModel)connectionResult.Data.SapModel;
            int unitRet = sapModel.SetPresentUnits(SAP2000v1.eUnits.N_mm_C);
            if (unitRet != 0)
            {
                return OperationResult.Failure("Failed to set SAP2000 present units to N-mm-C.");
            }

            int failCount = 0;
            var progress = BatchProgressWindow.RunWithProgress(inputs.Count, title, ctx =>
            {
                foreach (var input in inputs)
                {
                    if (ctx.IsCancellationRequested) break;

                    string sectionName = (string)input.GetType().GetProperty("SectionName").GetValue(input, null);
                    if (FrameSectionExists(sapModel, sectionName))
                    {
                        ctx.IncrementSkipped();
                        continue;
                    }

                    int ret = createSection(sapModel, input);
                    if (ret == 0)
                    {
                        ctx.IncrementRan();
                    }
                    else
                    {
                        failCount++;
                        ctx.IncrementSkipped();
                    }
                }
            });

            string msg = $"Created: {progress.RanCount}, Skipped: {progress.SkippedCount}, Failed: {failCount}";
            if (progress.WasCancelled) msg += " (Cancelled)";
            return OperationResult.Success(msg);
        }

        private bool FrameSectionExists(SAP2000v1.cSapModel sapModel, string sectionName)
        {
            if (string.IsNullOrWhiteSpace(sectionName))
            {
                return false;
            }

            SAP2000v1.eFramePropType propType = SAP2000v1.eFramePropType.I;
            int ret = sapModel.PropFrame.GetTypeOAPI(sectionName, ref propType);
            return ret == 0;
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

