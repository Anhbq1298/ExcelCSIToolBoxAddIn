using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using ExcelCSIToolBoxAddIn.Common.Results;
using ExcelCSIToolBoxAddIn.Core.Geometry;
using ExcelCSIToolBoxAddIn.UI.Views;
using SAP2000v1;

namespace ExcelCSIToolBoxAddIn.Infrastructure.Etabs
{
    /// <summary>
    /// SAP2000 adapter that exposes the same toolbox contract used by ETABS.
    /// The application/use-case layer stays shared; only the CSI API binding differs here.
    /// </summary>
    public class Sap2000ConnectionService : ICsiConnectionService
    {
        private const string Sap2000ComProgId = "CSI.SAP2000.API.SapObject";

        private CsiConnectionInfo _currentConnection;

        public string ProductName => "SAP2000";

        public OperationResult<CsiConnectionInfo> TryAttachToRunningInstance()
        {
            SAP2000v1.cHelper helper = new SAP2000v1.Helper();
            SAP2000v1.cOAPI sapObject = null;

            try
            {
                sapObject = helper.GetObject(Sap2000ComProgId);
                if (sapObject == null)
                {
                    _currentConnection = null;
                    return OperationResult<CsiConnectionInfo>.Failure("SAP2000 is not running.");
                }

                SAP2000v1.cSapModel sapModel = sapObject.SapModel;

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

                _currentConnection = new CsiConnectionInfo
                {
                    IsConnected = true,
                    ModelPath = modelPath,
                    ModelFileName = modelName,
                    ModelCurrentUnit = modelCurrentUnit,
                    CsiObject = sapObject,
                    SapModel = sapModel
                };

                return OperationResult<CsiConnectionInfo>.Success(_currentConnection);
            }
            catch
            {
                _currentConnection = null;
                return OperationResult<CsiConnectionInfo>.Failure("SAP2000 is not running.");
            }
        }

        public OperationResult<CsiConnectionInfo> GetCurrentConnection()
        {
            if (_currentConnection?.SapModel == null)
            {
                return OperationResult<CsiConnectionInfo>.Failure("No SAP2000 model is currently connected. Please click Attach.");
            }

            return OperationResult<CsiConnectionInfo>.Success(_currentConnection);
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
            return SelectObjectsByUniqueNames(uniqueNames, "point",
                (sapModel, name) => sapModel.PointObj.SetSelected(name, true, SAP2000v1.eItemType.Objects));
        }

        public OperationResult SelectFramesByUniqueNames(IReadOnlyList<string> uniqueNames)
        {
            return SelectObjectsByUniqueNames(uniqueNames, "frame",
                (sapModel, name) => sapModel.FrameObj.SetSelected(name, true, SAP2000v1.eItemType.Objects));
        }

        public OperationResult<CsiAddPointsResult> AddPointsByCartesian(IReadOnlyList<EtabsPointCartesianInput> pointInputs)
        {
            if (pointInputs == null || pointInputs.Count == 0)
            {
                return OperationResult<CsiAddPointsResult>.Failure("No valid rows were found in the selected range.");
            }

            var connectionResult = EnsureConnection();
            if (!connectionResult.IsSuccess || connectionResult.Data?.SapModel == null)
            {
                return OperationResult<CsiAddPointsResult>.Failure(connectionResult.Message);
            }

            try
            {
                SAP2000v1.cSapModel sapModel = (SAP2000v1.cSapModel)connectionResult.Data.SapModel;
                var failedRowMessages = new List<string>();
                var successCount = 0;

                BatchProgressWindow.RunWithProgress(pointInputs.Count, "Adding Points to Model...", ctx =>
                {
                    foreach (var pointInput in pointInputs)
                    {
                        if (ctx.IsCancellationRequested) break;

                        string assignedName = string.Empty;
                        string requestedUniqueName = string.IsNullOrWhiteSpace(pointInput.UniqueName) ? string.Empty : pointInput.UniqueName;

                        int addResult = sapModel.PointObj.AddCartesian(
                            pointInput.X,
                            pointInput.Y,
                            pointInput.Z,
                            ref assignedName,
                            requestedUniqueName,
                            "Global",
                            true,
                            0);

                        if (addResult != 0)
                        {
                            failedRowMessages.Add($"Row {pointInput.ExcelRowNumber}: SAP2000 API call PointObj.AddCartesian failed (return code {addResult}).");
                            ctx.IncrementSkipped();
                            continue;
                        }

                        successCount++;
                        ctx.IncrementRan();

                        if (!string.IsNullOrWhiteSpace(requestedUniqueName) &&
                            !string.Equals(assignedName, requestedUniqueName, StringComparison.OrdinalIgnoreCase))
                        {
                            failedRowMessages.Add($"Row {pointInput.ExcelRowNumber}: Point was created, but SAP2000 assigned UniqueName '{assignedName}' instead of requested '{requestedUniqueName}'.");
                        }
                    }
                });

                if (successCount > 0)
                {
                    var refreshResult = RefreshView(sapModel);
                    if (!refreshResult.IsSuccess)
                    {
                        return OperationResult<CsiAddPointsResult>.Failure(refreshResult.Message);
                    }
                }

                return OperationResult<CsiAddPointsResult>.Success(new CsiAddPointsResult
                {
                    AddedCount = successCount,
                    FailedRowMessages = failedRowMessages
                });
            }
            catch (COMException ex)
            {
                return OperationResult<CsiAddPointsResult>.Failure($"SAP2000 COM error while adding points by Cartesian coordinates: {ex.Message}");
            }
            catch (Exception ex)
            {
                return OperationResult<CsiAddPointsResult>.Failure($"SAP2000 add-by-Cartesian failed unexpectedly: {ex.Message}");
            }
        }

        public OperationResult<CsiAddFramesResult> AddFramesByCoordinates(IReadOnlyList<EtabsFrameByCoordInput> frameInputs)
        {
            if (frameInputs == null || frameInputs.Count == 0)
            {
                return OperationResult<CsiAddFramesResult>.Failure("No valid rows were found in the selected range.");
            }

            var connectionResult = EnsureConnection();
            if (!connectionResult.IsSuccess || connectionResult.Data?.SapModel == null)
            {
                return OperationResult<CsiAddFramesResult>.Failure(connectionResult.Message);
            }

            try
            {
                SAP2000v1.cSapModel sapModel = (SAP2000v1.cSapModel)connectionResult.Data.SapModel;
                return AddFrames(frameInputs.Count, "Adding Frames (by Coordinates)...", sapModel, ctx =>
                {
                    var failedRowMessages = new List<string>();
                    var successCount = 0;

                    foreach (var frameInput in frameInputs)
                    {
                        if (ctx.IsCancellationRequested) break;

                        string createdName = string.Empty;
                        string sectionName = string.IsNullOrWhiteSpace(frameInput.SectionName) ? "Default" : frameInput.SectionName;
                        string userName = string.IsNullOrWhiteSpace(frameInput.UniqueName) ? string.Empty : frameInput.UniqueName;

                        int addResult = sapModel.FrameObj.AddByCoord(
                            frameInput.Xi,
                            frameInput.Yi,
                            frameInput.Zi,
                            frameInput.Xj,
                            frameInput.Yj,
                            frameInput.Zj,
                            ref createdName,
                            sectionName,
                            userName,
                            "Global");

                        if (addResult != 0)
                        {
                            failedRowMessages.Add($"Row {frameInput.ExcelRowNumber}: SAP2000 API call FrameObj.AddByCoord failed (return code {addResult}).");
                            ctx.IncrementSkipped();
                            continue;
                        }

                        successCount++;
                        ctx.IncrementRan();
                    }

                    return new CsiAddFramesResult
                    {
                        AddedCount = successCount,
                        FailedRowMessages = failedRowMessages
                    };
                });
            }
            catch (Exception ex)
            {
                return OperationResult<CsiAddFramesResult>.Failure($"SAP2000 add-by-coordinates failed unexpectedly: {ex.Message}");
            }
        }

        public OperationResult<CsiAddFramesResult> AddFramesByPoint(IReadOnlyList<EtabsFrameByPointInput> frameInputs)
        {
            if (frameInputs == null || frameInputs.Count == 0)
            {
                return OperationResult<CsiAddFramesResult>.Failure("No valid rows were found in the selected range.");
            }

            var connectionResult = EnsureConnection();
            if (!connectionResult.IsSuccess || connectionResult.Data?.SapModel == null)
            {
                return OperationResult<CsiAddFramesResult>.Failure(connectionResult.Message);
            }

            try
            {
                SAP2000v1.cSapModel sapModel = (SAP2000v1.cSapModel)connectionResult.Data.SapModel;
                return AddFrames(frameInputs.Count, "Adding Frames (by Point Names)...", sapModel, ctx =>
                {
                    var failedRowMessages = new List<string>();
                    var successCount = 0;

                    foreach (var frameInput in frameInputs)
                    {
                        if (ctx.IsCancellationRequested) break;

                        string createdName = string.Empty;
                        string sectionName = string.IsNullOrWhiteSpace(frameInput.SectionName) ? "Default" : frameInput.SectionName;
                        string userName = string.IsNullOrWhiteSpace(frameInput.UniqueName) ? string.Empty : frameInput.UniqueName;

                        int addResult = sapModel.FrameObj.AddByPoint(
                            frameInput.Point1Name,
                            frameInput.Point2Name,
                            ref createdName,
                            sectionName,
                            userName);

                        if (addResult != 0)
                        {
                            failedRowMessages.Add($"Row {frameInput.ExcelRowNumber}: SAP2000 API call FrameObj.AddByPoint failed (return code {addResult}).");
                            ctx.IncrementSkipped();
                            continue;
                        }

                        successCount++;
                        ctx.IncrementRan();

                        if (!string.IsNullOrWhiteSpace(userName) &&
                            !string.Equals(createdName, userName, StringComparison.OrdinalIgnoreCase))
                        {
                            failedRowMessages.Add($"Row {frameInput.ExcelRowNumber}: Frame was created, but SAP2000 assigned UniqueName '{createdName}' instead of requested '{userName}'.");
                        }
                    }

                    return new CsiAddFramesResult
                    {
                        AddedCount = successCount,
                        FailedRowMessages = failedRowMessages
                    };
                });
            }
            catch (Exception ex)
            {
                return OperationResult<CsiAddFramesResult>.Failure($"SAP2000 add-by-point failed unexpectedly: {ex.Message}");
            }
        }

        public OperationResult<IReadOnlyList<CsiPointData>> GetSelectedPointsFromActiveModel()
        {
            var connectionResult = EnsureConnection();
            if (!connectionResult.IsSuccess || connectionResult.Data?.SapModel == null)
            {
                return OperationResult<IReadOnlyList<CsiPointData>>.Failure(connectionResult.Message);
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
                    return OperationResult<IReadOnlyList<CsiPointData>>.Failure("Failed to read selected objects from SAP2000.");
                }

                var points = new List<CsiPointData>();
                for (int i = 0; i < numberItems; i++)
                {
                    if (objectTypes == null || objectNames == null || i >= objectTypes.Length || i >= objectNames.Length)
                    {
                        continue;
                    }

                    if (objectTypes[i] != CsiObjectTypeIds.Point || string.IsNullOrWhiteSpace(objectNames[i]))
                    {
                        continue;
                    }

                    double x = 0;
                    double y = 0;
                    double z = 0;
                    int pointResult = sapModel.PointObj.GetCoordCartesian(objectNames[i], ref x, ref y, ref z, "Global");
                    if (pointResult == 0)
                    {
                        points.Add(new CsiPointData
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
                    return OperationResult<IReadOnlyList<CsiPointData>>.Failure("No point objects are selected in SAP2000.");
                }

                return OperationResult<IReadOnlyList<CsiPointData>>.Success(points);
            }
            catch
            {
                return OperationResult<IReadOnlyList<CsiPointData>>.Failure("Unable to read selected SAP2000 points.");
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
                    if (objectTypes[i] == CsiObjectTypeIds.Frame && !string.IsNullOrWhiteSpace(frameUniqueName))
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

        public OperationResult AddSteelISections(IReadOnlyList<EtabsSteelISectionInput> inputs)
        {
            return CreateSections(inputs, "Creating Steel I-Sections...", (sapModel, input) =>
                sapModel.PropFrame.SetISection(input.SectionName, input.MaterialName, input.H, input.B, input.Tf, input.Tw, input.B, input.Tf, -1, "", ""));
        }

        public OperationResult AddSteelChannelSections(IReadOnlyList<EtabsSteelChannelSectionInput> inputs)
        {
            return CreateSections(inputs, "Creating Steel Channel Sections...", (sapModel, input) =>
                sapModel.PropFrame.SetChannel(input.SectionName, input.MaterialName, input.H, input.B, input.Tf, input.Tw, -1, "", ""));
        }

        public OperationResult AddSteelAngleSections(IReadOnlyList<EtabsSteelAngleSectionInput> inputs)
        {
            return CreateSections(inputs, "Creating Steel Angle Sections...", (sapModel, input) =>
                sapModel.PropFrame.SetAngle(input.SectionName, input.MaterialName, input.H, input.B, input.Tf, input.Tw, -1, "", ""));
        }

        public OperationResult AddSteelPipeSections(IReadOnlyList<EtabsSteelPipeSectionInput> inputs)
        {
            return CreateSections(inputs, "Creating Steel Pipe Sections...", (sapModel, input) =>
                sapModel.PropFrame.SetPipe(input.SectionName, input.MaterialName, input.OutsideDiameter, input.WallThickness, -1, "", ""));
        }

        public OperationResult AddSteelTubeSections(IReadOnlyList<EtabsSteelTubeSectionInput> inputs)
        {
            return CreateSections(inputs, "Creating Steel Tube Sections...", (sapModel, input) =>
                sapModel.PropFrame.SetTube_1(input.SectionName, input.MaterialName, input.H, input.B, input.T, input.T, 0.000000001, -1, "", ""));
        }

        public OperationResult AddConcreteRectangleSections(IReadOnlyList<EtabsConcreteRectangleSectionInput> inputs)
        {
            return CreateSections(inputs, "Creating Concrete Rectangle Sections...", (sapModel, input) =>
                sapModel.PropFrame.SetRectangle(input.SectionName, input.MaterialName, input.H, input.B, -1, "", ""));
        }

        public OperationResult AddConcreteCircleSections(IReadOnlyList<EtabsConcreteCircleSectionInput> inputs)
        {
            return CreateSections(inputs, "Creating Concrete Circle Sections...", (sapModel, input) =>
                sapModel.PropFrame.SetCircle(input.SectionName, input.MaterialName, input.D, -1, "", ""));
        }

        public OperationResult CreateShellAreasFromSelectedFrames(
            string propertyName,
            ShellCreationTolerances tolerances)
        {
            tolerances = tolerances ?? new ShellCreationTolerances();
            propertyName = string.IsNullOrWhiteSpace(propertyName) ? "Default" : propertyName.Trim();

            var connectionResult = EnsureConnection();
            if (!connectionResult.IsSuccess || connectionResult.Data?.SapModel == null)
            {
                return OperationResult.Failure(connectionResult.Message);
            }

            try
            {
                SAP2000v1.cSapModel sapModel = (SAP2000v1.cSapModel)connectionResult.Data.SapModel;

                int unitRet = sapModel.SetPresentUnits(SAP2000v1.eUnits.kN_m_C);
                if (unitRet != 0)
                {
                    return OperationResult.Failure("Failed to set SAP2000 present units to kN-m-C.");
                }

                var framesResult = ReadSelectedFrameGeometries(sapModel);
                if (!framesResult.IsSuccess)
                {
                    return OperationResult.Failure(framesResult.Message);
                }

                var frameGeometries = framesResult.Data;
                if (frameGeometries == null || frameGeometries.Count == 0)
                {
                    return OperationResult.Failure("No frame objects are currently selected in SAP2000.");
                }

                var faceBuildResult = ShellFaceBuilder.BuildCandidateFaces(frameGeometries, tolerances);
                if (faceBuildResult.EnrichedRealEdgeCount == 0)
                {
                    return OperationResult.Failure("No valid enriched frame graph was found from the current selection.");
                }

                if (faceBuildResult.FaceLoops == null || faceBuildResult.FaceLoops.Count == 0)
                {
                    return OperationResult.Failure("No closed shell faces were extracted from the selected frames.");
                }

                var acceptedFaces = new List<IReadOnlyList<string>>();
                var createdCount = 0;
                var skippedCount = 0;
                var shellFaceCandidates = BuildShellFaceCandidates(
                    faceBuildResult.FaceLoops,
                    faceBuildResult.PointCoordinates,
                    tolerances,
                    ref skippedCount);

                var progress = BatchProgressWindow.RunWithProgress(
                    shellFaceCandidates.Count,
                    "Creating Shell Areas From Selected Frames...",
                    ctx =>
                    {
                        foreach (var candidate in shellFaceCandidates)
                        {
                            if (ctx.IsCancellationRequested) break;

                            string rejectReason;
                            if (!ShellFaceBuilder.ValidateFaceLoop(
                                    candidate.OrderedLoop,
                                    faceBuildResult.PointCoordinates,
                                    acceptedFaces,
                                    tolerances,
                                    out rejectReason))
                            {
                                skippedCount++;
                                ctx.IncrementSkipped();
                                continue;
                            }

                            var createdForLoop = CreateAreaFromLoop(
                                sapModel,
                                candidate.OrderedLoop,
                                faceBuildResult.PointCoordinates,
                                propertyName,
                                acceptedFaces,
                                tolerances);

                            if (createdForLoop > 0)
                            {
                                createdCount += createdForLoop;
                                ctx.IncrementRan();
                            }
                            else
                            {
                                skippedCount++;
                                ctx.IncrementSkipped();
                            }
                        }
                    });

                var refreshResult = RefreshView(sapModel);
                if (!refreshResult.IsSuccess)
                {
                    return refreshResult;
                }

                var message = "Done." + Environment.NewLine +
                              $"Created: {createdCount}" + Environment.NewLine +
                              $"Skipped: {skippedCount}";

                if (progress.WasCancelled)
                {
                    message += Environment.NewLine + "Operation cancelled before all faces were processed.";
                }

                return OperationResult.Success(message);
            }
            catch (COMException ex)
            {
                return OperationResult.Failure($"SAP2000 COM error while creating shell areas from selected frames: {ex.Message}");
            }
            catch (Exception ex)
            {
                return OperationResult.Failure($"Failed to create shell areas from selected frames: {ex.Message}");
            }
        }

        private OperationResult SelectObjectsByUniqueNames(
            IReadOnlyList<string> uniqueNames,
            string objectTypeName,
            Func<SAP2000v1.cSapModel, string, int> setSelected)
        {
            if (uniqueNames == null || uniqueNames.Count == 0)
            {
                return OperationResult.Failure("The selected Excel range does not contain any non-empty values.");
            }

            var orderedUniqueNames = GetOrderedDistinctNames(uniqueNames);
            if (orderedUniqueNames.Count == 0)
            {
                return OperationResult.Failure("The selected Excel range does not contain any non-empty values.");
            }

            var connectionResult = EnsureConnection();
            if (!connectionResult.IsSuccess || connectionResult.Data?.SapModel == null)
            {
                return OperationResult.Failure(connectionResult.Message);
            }

            try
            {
                SAP2000v1.cSapModel sapModel = (SAP2000v1.cSapModel)connectionResult.Data.SapModel;
                int clearSelectionResult = sapModel.SelectObj.ClearSelection();
                if (clearSelectionResult != 0)
                {
                    return OperationResult.Failure($"Failed to clear SAP2000 selection before selecting {objectTypeName}s by UniqueName.");
                }

                var unresolved = new List<string>();
                var selectedCount = 0;

                BatchProgressWindow.RunWithProgress(orderedUniqueNames.Count, $"Selecting {objectTypeName}s...", ctx =>
                {
                    foreach (var uniqueName in orderedUniqueNames)
                    {
                        if (ctx.IsCancellationRequested) break;

                        int result = setSelected(sapModel, uniqueName);
                        if (result == 0)
                        {
                            selectedCount++;
                            ctx.IncrementRan();
                        }
                        else
                        {
                            unresolved.Add(uniqueName);
                            ctx.IncrementSkipped();
                        }
                    }
                });

                var message = $"Selected {selectedCount} {objectTypeName}(s) by UniqueName.";
                if (unresolved.Count > 0)
                {
                    message += $" Not found: {string.Join(", ", unresolved)}.";
                }

                var refreshResult = RefreshView(sapModel);
                if (!refreshResult.IsSuccess)
                {
                    return refreshResult;
                }

                return OperationResult.Success(message);
            }
            catch
            {
                return OperationResult.Failure($"Failed to select SAP2000 {objectTypeName}s by UniqueName.");
            }
        }

        private OperationResult<CsiConnectionInfo> EnsureConnection()
        {
            var connectionResult = GetCurrentConnection();
            if (!connectionResult.IsSuccess || connectionResult.Data?.SapModel == null)
            {
                connectionResult = TryAttachToRunningInstance();
            }

            return connectionResult;
        }

        private OperationResult<CsiAddFramesResult> AddFrames(
            int totalItems,
            string title,
            SAP2000v1.cSapModel sapModel,
            Func<BatchProgressContext, CsiAddFramesResult> addAction)
        {
            CsiAddFramesResult result = null;
            BatchProgressWindow.RunWithProgress(totalItems, title, ctx => result = addAction(ctx));

            if (result?.AddedCount > 0)
            {
                var refreshResult = RefreshView(sapModel);
                if (!refreshResult.IsSuccess)
                {
                    return OperationResult<CsiAddFramesResult>.Failure(refreshResult.Message);
                }
            }

            return OperationResult<CsiAddFramesResult>.Success(result ?? new CsiAddFramesResult());
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

        private OperationResult<IReadOnlyList<ShellFrameGeometry>> ReadSelectedFrameGeometries(SAP2000v1.cSapModel sapModel)
        {
            int numberItems = 0;
            int[] objectTypes = null;
            string[] objectNames = null;
            int selectedResult = sapModel.SelectObj.GetSelected(ref numberItems, ref objectTypes, ref objectNames);

            if (selectedResult != 0)
            {
                return OperationResult<IReadOnlyList<ShellFrameGeometry>>.Failure("Failed to read selected objects from SAP2000.");
            }

            if (numberItems <= 0 || objectTypes == null || objectNames == null)
            {
                return OperationResult<IReadOnlyList<ShellFrameGeometry>>.Failure("No frame objects are currently selected in SAP2000.");
            }

            var frames = new List<ShellFrameGeometry>();
            var seenFrames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            for (int i = 0; i < numberItems; i++)
            {
                if (i >= objectTypes.Length || i >= objectNames.Length)
                {
                    continue;
                }

                var frameName = objectNames[i];
                if (objectTypes[i] != CsiObjectTypeIds.Frame ||
                    string.IsNullOrWhiteSpace(frameName) ||
                    seenFrames.Contains(frameName))
                {
                    continue;
                }

                string p1 = string.Empty;
                string p2 = string.Empty;
                int framePointsResult = sapModel.FrameObj.GetPoints(frameName, ref p1, ref p2);
                if (framePointsResult != 0 || string.IsNullOrWhiteSpace(p1) || string.IsNullOrWhiteSpace(p2))
                {
                    continue;
                }

                double x1 = 0;
                double y1 = 0;
                double z1 = 0;
                double x2 = 0;
                double y2 = 0;
                double z2 = 0;

                int p1Result = sapModel.PointObj.GetCoordCartesian(p1, ref x1, ref y1, ref z1, "Global");
                int p2Result = sapModel.PointObj.GetCoordCartesian(p2, ref x2, ref y2, ref z2, "Global");
                if (p1Result != 0 || p2Result != 0)
                {
                    continue;
                }

                seenFrames.Add(frameName);
                frames.Add(new ShellFrameGeometry
                {
                    FrameName = frameName,
                    StartPointName = p1,
                    EndPointName = p2,
                    StartPoint = new ShellPoint3D(x1, y1, z1),
                    EndPoint = new ShellPoint3D(x2, y2, z2)
                });
            }

            if (frames.Count == 0)
            {
                return OperationResult<IReadOnlyList<ShellFrameGeometry>>.Failure("No valid frame geometry could be read from the current SAP2000 selection.");
            }

            return OperationResult<IReadOnlyList<ShellFrameGeometry>>.Success(frames);
        }

        private static List<ShellFaceCandidate> BuildShellFaceCandidates(
            IReadOnlyList<string[]> rawFaceLoops,
            IReadOnlyDictionary<string, ShellPoint3D> pointCoords,
            ShellCreationTolerances tolerances,
            ref int skippedCount)
        {
            var candidates = new List<ShellFaceCandidate>();
            var emptyAcceptedFaces = new List<IReadOnlyList<string>>();

            foreach (var rawLoop in rawFaceLoops)
            {
                var cleanLoop = ShellFaceBuilder.CleanLoopBoundaryXY(rawLoop, pointCoords, tolerances);
                if (cleanLoop == null || cleanLoop.Length < 3)
                {
                    skippedCount++;
                    continue;
                }

                var orderedLoop = ShellFaceBuilder.OrderLoopUpward(cleanLoop, pointCoords);
                string rejectReason;
                if (!ShellFaceBuilder.ValidateFaceLoop(
                        orderedLoop,
                        pointCoords,
                        emptyAcceptedFaces,
                        tolerances,
                        out rejectReason))
                {
                    skippedCount++;
                    continue;
                }

                candidates.Add(new ShellFaceCandidate
                {
                    OrderedLoop = orderedLoop,
                    Area = Math.Abs(ShellFaceBuilder.GetPolygonAreaXY(orderedLoop, pointCoords))
                });
            }

            return candidates
                .OrderBy(candidate => GetShellLoopPriority(candidate.OrderedLoop.Length))
                .ThenBy(candidate => candidate.OrderedLoop.Length)
                .ThenBy(candidate => candidate.Area)
                .ToList();
        }

        private static int GetShellLoopPriority(int nodeCount)
        {
            if (nodeCount == 4) return 0;
            if (nodeCount == 3) return 1;
            return 2;
        }

        private static int CreateAreaFromLoop(
            SAP2000v1.cSapModel sapModel,
            IReadOnlyList<string> loopPts,
            IReadOnlyDictionary<string, ShellPoint3D> pointCoords,
            string propName,
            List<IReadOnlyList<string>> acceptedFaces,
            ShellCreationTolerances tolerances)
        {
            if (loopPts == null)
            {
                return 0;
            }

            if (loopPts.Count == 3)
            {
                if (AddAreaByNodeCoordinates(sapModel, loopPts, pointCoords, propName))
                {
                    acceptedFaces.Add(loopPts.ToArray());
                    return 1;
                }

                return 0;
            }

            if (loopPts.Count == 4)
            {
                if (AddAreaByNodeCoordinates(sapModel, loopPts, pointCoords, propName))
                {
                    acceptedFaces.Add(loopPts.ToArray());
                    return 1;
                }

                return SplitQuadAndCreateTwoTriangles(sapModel, loopPts, pointCoords, propName, acceptedFaces, tolerances);
            }

            return 0;
        }

        private static bool AddAreaByNodeCoordinates(
            SAP2000v1.cSapModel sapModel,
            IReadOnlyList<string> nodeIds,
            IReadOnlyDictionary<string, ShellPoint3D> pointCoords,
            string propName)
        {
            if (nodeIds == null || (nodeIds.Count != 3 && nodeIds.Count != 4))
            {
                return false;
            }

            var x = new double[nodeIds.Count];
            var y = new double[nodeIds.Count];
            var z = new double[nodeIds.Count];

            for (int i = 0; i < nodeIds.Count; i++)
            {
                ShellPoint3D point;
                if (!pointCoords.TryGetValue(nodeIds[i], out point))
                {
                    return false;
                }

                x[i] = point.X;
                y[i] = point.Y;
                z[i] = point.Z;
            }

            string areaName = string.Empty;
            int addResult = sapModel.AreaObj.AddByCoord(nodeIds.Count, ref x, ref y, ref z, ref areaName, propName, string.Empty, "Global");
            return addResult == 0;
        }

        private static int SplitQuadAndCreateTwoTriangles(
            SAP2000v1.cSapModel sapModel,
            IReadOnlyList<string> quadPts,
            IReadOnlyDictionary<string, ShellPoint3D> pointCoords,
            string propName,
            List<IReadOnlyList<string>> acceptedFaces,
            ShellCreationTolerances tolerances)
        {
            var acLen = ShellFaceBuilder.Distance3D(pointCoords[quadPts[0]], pointCoords[quadPts[2]]);
            var bdLen = ShellFaceBuilder.Distance3D(pointCoords[quadPts[1]], pointCoords[quadPts[3]]);

            string[] tri1;
            string[] tri2;

            if (acLen <= bdLen)
            {
                tri1 = new[] { quadPts[0], quadPts[1], quadPts[2] };
                tri2 = new[] { quadPts[0], quadPts[2], quadPts[3] };
            }
            else
            {
                tri1 = new[] { quadPts[0], quadPts[1], quadPts[3] };
                tri2 = new[] { quadPts[1], quadPts[2], quadPts[3] };
            }

            var tri1Up = ShellFaceBuilder.OrderLoopUpward(tri1, pointCoords);
            var tri2Up = ShellFaceBuilder.OrderLoopUpward(tri2, pointCoords);

            if (ShellFaceBuilder.IsDegenerateTriangle(tri1Up, pointCoords) ||
                ShellFaceBuilder.IsDegenerateTriangle(tri2Up, pointCoords))
            {
                return 0;
            }

            if (ShellFaceBuilder.OverlapsAcceptedFacesXY(tri1Up, acceptedFaces, pointCoords, tolerances.PointTolerance) ||
                ShellFaceBuilder.OverlapsAcceptedFacesXY(tri2Up, acceptedFaces, pointCoords, tolerances.PointTolerance))
            {
                return 0;
            }

            if (!AddAreaByNodeCoordinates(sapModel, tri1Up, pointCoords, propName))
            {
                return 0;
            }

            acceptedFaces.Add(tri1Up);

            if (!AddAreaByNodeCoordinates(sapModel, tri2Up, pointCoords, propName))
            {
                return 1;
            }

            acceptedFaces.Add(tri2Up);
            return 2;
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

        private static IReadOnlyList<string> GetOrderedDistinctNames(IReadOnlyList<string> names)
        {
            var uniqueNames = new List<string>();
            var seenNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (var rawName in names)
            {
                var name = string.IsNullOrWhiteSpace(rawName) ? null : rawName.Trim();
                if (string.IsNullOrWhiteSpace(name) || seenNames.Contains(name))
                {
                    continue;
                }

                seenNames.Add(name);
                uniqueNames.Add(name);
            }

            return uniqueNames;
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

        private class ShellFaceCandidate
        {
            public string[] OrderedLoop { get; set; }
            public double Area { get; set; }
        }
    }
}

