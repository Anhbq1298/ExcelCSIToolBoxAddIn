using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using ExcelCSIToolBoxAddIn.Adapters;
using ExcelCSIToolBoxAddIn.Common.Results;
using ExcelCSIToolBoxAddIn.Core.Geometry;
using ExcelCSIToolBoxAddIn.Infrastructure.CSISapModel;
using ExcelCSIToolBoxAddIn.UI.Views;

namespace ExcelCSIToolBoxAddIn.Infrastructure.Etabs
{
    /// <summary>
    /// Infrastructure service that safely attaches to a running ETABS instance.
    /// Stores the latest attached instance so ETABS commands can reuse the same SapModel.
    /// </summary>
    public class EtabsConnectionService : ICSISapModelConnectionService
    {
        private readonly ICsiModelAdapter _modelAdapter;
        private CSISapModelConnectionInfo _currentConnection;

        public EtabsConnectionService()
            : this(new EtabsModelAdapter())
        {
        }

        public EtabsConnectionService(ICsiModelAdapter modelAdapter)
        {
            _modelAdapter = modelAdapter ?? throw new ArgumentNullException(nameof(modelAdapter));
        }

        public string ProductName => "ETABS";

        public OperationResult<CSISapModelConnectionInfo> TryAttachToRunningInstance()
        {
            var attachResult = _modelAdapter.AttachToRunningInstance();
            if (!attachResult.IsSuccess)
            {
                _currentConnection = null;
                return OperationResult<CSISapModelConnectionInfo>.Failure(attachResult.Message);
            }

            var etabsObject = attachResult.ApplicationObject as ETABSv1.cOAPI;
            var sapModel = attachResult.SapModel as ETABSv1.cSapModel;
            if (etabsObject == null || sapModel == null)
            {
                _currentConnection = null;
                return OperationResult<CSISapModelConnectionInfo>.Failure("The attached ETABS instance is invalid. Please reattach and try again.");
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
                    // Keep attach successful when model metadata retrieval fails.
                }

                string modelCurrentUnit = "Units unavailable";
                try
                {
                    ETABSv1.eForce forceUnits = ETABSv1.eForce.N;
                    ETABSv1.eLength lengthUnits = ETABSv1.eLength.m;
                    ETABSv1.eTemperature temperatureUnits = ETABSv1.eTemperature.C;

                    int getUnitsResult = sapModel.GetDatabaseUnits_2(ref forceUnits, ref lengthUnits, ref temperatureUnits);

                    if (getUnitsResult == 0)
                    {
                        modelCurrentUnit = EtabsUnitFormatter.FormatDatabaseUnits(forceUnits, lengthUnits, temperatureUnits);
                    }
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
                    CsiObject = etabsObject,
                    SapModel = sapModel
                };

                return OperationResult<CSISapModelConnectionInfo>.Success(_currentConnection);
            }
            catch
            {
                _currentConnection = null;
                return OperationResult<CSISapModelConnectionInfo>.Failure("Failed to attach to the running ETABS instance.");
            }
        }

        public OperationResult<CSISapModelConnectionInfo> GetCurrentConnection()
        {
            if (_currentConnection?.SapModel == null)
            {
                return OperationResult<CSISapModelConnectionInfo>.Failure("No ETABS model is currently connected. Please click 'Attach to Running ETABS'.");
            }

            return OperationResult<CSISapModelConnectionInfo>.Success(_currentConnection);
        }

        public OperationResult CloseCurrentInstance()
        {
            if (_currentConnection?.CsiObject == null)
            {
                return OperationResult.Failure("No running ETABS instance is currently attached.");
            }

            try
            {
                var etabsApplication = _currentConnection.CsiObject as ETABSv1.cOAPI;
                if (etabsApplication == null)
                {
                    return OperationResult.Failure("The attached ETABS instance is invalid. Please reattach and try again.");
                }

                int result = etabsApplication.ApplicationExit(false);
                if (result != 0)
                {
                    return OperationResult.Failure($"ETABS failed to close the attached instance (ApplicationExit returned {result}).");
                }

                ResetCurrentConnection();
                return OperationResult.Success("Successfully closed the attached ETABS instance.");
            }
            catch (COMException ex)
            {
                return OperationResult.Failure($"ETABS COM error while closing attached instance: {ex.Message}");
            }
            catch
            {
                return OperationResult.Failure("Failed to close the attached ETABS instance.");
            }
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
                // Ignored to avoid masking the primary ETABS operation result.
            }
        }

        public OperationResult SelectPointsByUniqueNames(IReadOnlyList<string> uniqueNames)
        {
            return CSISapModelOperationRunner.SelectObjectsByUniqueNames(
                uniqueNames,
                "point",
                ProductName,
                EnsureConnection,
                sapModel => (ETABSv1.cSapModel)sapModel,
                sapModel => sapModel.SelectObj.ClearSelection(),
                (sapModel, name) => sapModel.PointObj.SetSelected(name, true, ETABSv1.eItemType.Objects),
                RefreshView);
        }

        public OperationResult SelectFramesByUniqueNames(IReadOnlyList<string> uniqueNames)
        {
            return CSISapModelOperationRunner.SelectObjectsByUniqueNames(
                uniqueNames,
                "frame",
                ProductName,
                EnsureConnection,
                sapModel => (ETABSv1.cSapModel)sapModel,
                sapModel => sapModel.SelectObj.ClearSelection(),
                (sapModel, name) => sapModel.FrameObj.SetSelected(name, true, ETABSv1.eItemType.Objects),
                RefreshView);
        }

        public OperationResult<CSISapModelAddPointsResult> AddPointsByCartesian(IReadOnlyList<CSISapModelPointCartesianInput> pointInputs)
        {
            return CSISapModelOperationRunner.AddPointsByCartesian(
                pointInputs,
                ProductName,
                EnsureConnection,
                sapModel => (ETABSv1.cSapModel)sapModel,
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
        }

        public OperationResult<CSISapModelAddFramesResult> AddFramesByCoordinates(IReadOnlyList<CSISapModelFrameByCoordInput> frameInputs)
        {
            return CSISapModelOperationRunner.AddFramesByCoordinates(
                frameInputs,
                ProductName,
                EnsureConnection,
                sapModel => (ETABSv1.cSapModel)sapModel,
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

        public OperationResult<CSISapModelAddFramesResult> AddFramesByPoint(IReadOnlyList<CSISapModelFrameByPointInput> frameInputs)
        {
            return CSISapModelOperationRunner.AddFramesByPoint(
                frameInputs,
                ProductName,
                EnsureConnection,
                sapModel => (ETABSv1.cSapModel)sapModel,
                (ETABSv1.cSapModel sapModel, CSISapModelFrameByPointInput frameInput, ref string createdName, string sectionName, string userName) =>
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
                ETABSv1.cSapModel sapModel = (ETABSv1.cSapModel)connectionResult.Data.SapModel;

                int numberItems = 0;
                int[] objectTypes = null;
                string[] objectNames = null;
                int selectedResult = sapModel.SelectObj.GetSelected(ref numberItems, ref objectTypes, ref objectNames);

                if (selectedResult != 0)
                {
                    return OperationResult<IReadOnlyList<CSISapModelPointData>>.Failure("Failed to read selected objects from ETABS.");
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
                    string pointLabel = string.Empty;
                    string pointStory = string.Empty;
                    int pointLabelResult = sapModel.PointObj.GetLabelFromName(objectNames[i], ref pointLabel, ref pointStory);

                    if (pointResult == 0)
                    {
                        points.Add(new CSISapModelPointData
                        {
                            PointUniqueName = objectNames[i],
                            PointLabel = pointLabelResult == 0 && !string.IsNullOrWhiteSpace(pointLabel)
                                ? pointLabel
                                : "(Unresolved)",
                            X = x,
                            Y = y,
                            Z = z
                        });
                    }
                }

                if (points.Count == 0)
                {
                    return OperationResult<IReadOnlyList<CSISapModelPointData>>.Failure("No point objects are selected in ETABS.");
                }

                return OperationResult<IReadOnlyList<CSISapModelPointData>>.Success(points);
            }
            catch
            {
                return OperationResult<IReadOnlyList<CSISapModelPointData>>.Failure("Unable to read selected ETABS points.");
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
                ETABSv1.cSapModel sapModel = (ETABSv1.cSapModel)connectionResult.Data.SapModel;

                int numberItems = 0;
                int[] objectTypes = null;
                string[] objectNames = null;
                int selectedResult = sapModel.SelectObj.GetSelected(ref numberItems, ref objectTypes, ref objectNames);
                if (selectedResult != 0)
                {
                    return OperationResult<IReadOnlyList<string>>.Failure("Failed to read selected objects from ETABS.");
                }

                if (numberItems < 0)
                {
                    return OperationResult<IReadOnlyList<string>>.Failure("Failed to read selected objects from ETABS.");
                }

                if (numberItems > 0 &&
                    (objectTypes == null ||
                     objectNames == null ||
                     objectTypes.Length < numberItems ||
                     objectNames.Length < numberItems))
                {
                    return OperationResult<IReadOnlyList<string>>.Failure("Selected object data from ETABS is inconsistent.");
                }

                var frameUniqueNames = new List<string>();

                for (int i = 0; i < numberItems; i++)
                {
                    var frameUniqueName = objectNames[i];
                    if (objectTypes[i] != CSISapModelObjectTypeIds.Frame || string.IsNullOrWhiteSpace(frameUniqueName))
                    {
                        continue;
                    }

                    frameUniqueNames.Add(frameUniqueName);
                }

                if (frameUniqueNames.Count == 0)
                {
                    return OperationResult<IReadOnlyList<string>>.Failure("No frame objects are selected in ETABS.");
                }

                return OperationResult<IReadOnlyList<string>>.Success(frameUniqueNames);
            }
            catch
            {
                return OperationResult<IReadOnlyList<string>>.Failure("Unable to read selected ETABS frames.");
            }
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

        public OperationResult AddSteelISections(IReadOnlyList<CSISapModelSteelISectionInput> inputs)
        {
            ETABSv1.cSapModel sapModel = _currentConnection?.SapModel as ETABSv1.cSapModel;
            if (sapModel == null)
            {
                return OperationResult.Failure("No active ETABS model found.");
            }

            int unitRet = sapModel.SetPresentUnits(ETABSv1.eUnits.N_mm_C);
            if (unitRet != 0)
            {
                return OperationResult.Failure("Failed to set ETABS present units to N-mm-C.");
            }

            int failCount = 0;

            var progress = BatchProgressWindow.RunWithProgress(inputs.Count, "Creating Steel I-Sections...", (ctx) =>
            {
                foreach (var input in inputs)
                {
                    if (ctx.IsCancellationRequested) break;

                    if (FrameSectionExists(sapModel, input.SectionName))
                    {
                        ctx.IncrementSkipped();
                        continue;
                    }

                    int ret = sapModel.PropFrame.SetISection(
                        input.SectionName,
                        input.MaterialName,
                        input.H,
                        input.B,
                        input.Tf,
                        input.Tw,
                        input.B,
                        input.Tf);

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

        public OperationResult AddSteelChannelSections(IReadOnlyList<CSISapModelSteelChannelSectionInput> inputs)
        {
            ETABSv1.cSapModel sapModel = _currentConnection?.SapModel as ETABSv1.cSapModel;
            if (sapModel == null)
            {
                return OperationResult.Failure("No active ETABS model found.");
            }

            int unitRet = sapModel.SetPresentUnits(ETABSv1.eUnits.N_mm_C);
            if (unitRet != 0)
            {
                return OperationResult.Failure("Failed to set ETABS present units to N-mm-C.");
            }

            int failCount = 0;

            var progress = BatchProgressWindow.RunWithProgress(inputs.Count, "Creating Steel Channel Sections...", (ctx) =>
            {
                foreach (var input in inputs)
                {
                    if (ctx.IsCancellationRequested) break;

                    if (FrameSectionExists(sapModel, input.SectionName))
                    {
                        ctx.IncrementSkipped();
                        continue;
                    }

                    int ret = sapModel.PropFrame.SetChannel(
                        input.SectionName,
                        input.MaterialName,
                        input.H,
                        input.B,
                        input.Tf,
                        input.Tw);

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

        public OperationResult AddSteelAngleSections(IReadOnlyList<CSISapModelSteelAngleSectionInput> inputs)
        {
            ETABSv1.cSapModel sapModel = _currentConnection?.SapModel as ETABSv1.cSapModel;
            if (sapModel == null)
            {
                return OperationResult.Failure("No active ETABS model found.");
            }

            int unitRet = sapModel.SetPresentUnits(ETABSv1.eUnits.N_mm_C);
            if (unitRet != 0)
            {
                return OperationResult.Failure("Failed to set ETABS present units to N-mm-C.");
            }

            int failCount = 0;

            var progress = BatchProgressWindow.RunWithProgress(inputs.Count, "Creating Steel Angle Sections...", (ctx) =>
            {
                foreach (var input in inputs)
                {
                    if (ctx.IsCancellationRequested) break;

                    if (FrameSectionExists(sapModel, input.SectionName))
                    {
                        ctx.IncrementSkipped();
                        continue;
                    }

                    int ret = sapModel.PropFrame.SetAngle(
                        input.SectionName,
                        input.MaterialName,
                        input.H,
                        input.B,
                        input.Tf,
                        input.Tw);

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

        public OperationResult AddSteelPipeSections(IReadOnlyList<CSISapModelSteelPipeSectionInput> inputs)
        {
            ETABSv1.cSapModel sapModel = _currentConnection?.SapModel as ETABSv1.cSapModel;
            if (sapModel == null)
            {
                return OperationResult.Failure("No active ETABS model found.");
            }

            int unitRet = sapModel.SetPresentUnits(ETABSv1.eUnits.N_mm_C);
            if (unitRet != 0)
            {
                return OperationResult.Failure("Failed to set ETABS present units to N-mm-C.");
            }

            int failCount = 0;

            var progress = BatchProgressWindow.RunWithProgress(inputs.Count, "Creating Steel Pipe Sections...", (ctx) =>
            {
                foreach (var input in inputs)
                {
                    if (ctx.IsCancellationRequested) break;

                    if (FrameSectionExists(sapModel, input.SectionName))
                    {
                        ctx.IncrementSkipped();
                        continue;
                    }

                    int ret = sapModel.PropFrame.SetPipe(
                        input.SectionName,
                        input.MaterialName,
                        input.OutsideDiameter,
                        input.WallThickness);

                    if (ret == 0)
                        ctx.IncrementRan();
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

        public OperationResult AddSteelTubeSections(IReadOnlyList<CSISapModelSteelTubeSectionInput> inputs)
        {
            ETABSv1.cSapModel sapModel = _currentConnection?.SapModel as ETABSv1.cSapModel;
            if (sapModel == null)
            {
                return OperationResult.Failure("No active ETABS model found.");
            }

            int unitRet = sapModel.SetPresentUnits(ETABSv1.eUnits.N_mm_C);
            if (unitRet != 0)
            {
                return OperationResult.Failure("Failed to set ETABS present units to N-mm-C.");
            }

            int failCount = 0;

            var progress = BatchProgressWindow.RunWithProgress(inputs.Count, "Creating Steel Tube Sections...", (ctx) =>
            {
                foreach (var input in inputs)
                {
                    if (ctx.IsCancellationRequested) break;

                    if (FrameSectionExists(sapModel, input.SectionName))
                    {
                        ctx.IncrementSkipped();
                        continue;
                    }

                    int ret = sapModel.PropFrame.SetTube_1(
                        input.SectionName,
                        input.MaterialName,
                        input.H,
                        input.B,
                        input.T,
                        input.T,
                        0.000000001,
                        -1,
                        "",
                        "Default");

                    if (ret == 0)
                        ctx.IncrementRan();
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

        public OperationResult AddConcreteRectangleSections(IReadOnlyList<CSISapModelConcreteRectangleSectionInput> inputs)
        {
            ETABSv1.cSapModel sapModel = _currentConnection?.SapModel as ETABSv1.cSapModel;
            if (sapModel == null) return OperationResult.Failure("No active ETABS model found.");

            int unitRet = sapModel.SetPresentUnits(ETABSv1.eUnits.N_mm_C);
            if (unitRet != 0) return OperationResult.Failure("Failed to set ETABS present units to N-mm-C.");

            int failCount = 0;

            var progress = BatchProgressWindow.RunWithProgress(inputs.Count, "Creating Concrete Rectangle Sections...", (ctx) =>
            {
                foreach (var input in inputs)
                {
                    if (ctx.IsCancellationRequested) break;

                    if (FrameSectionExists(sapModel, input.SectionName))
                    {
                        ctx.IncrementSkipped();
                        continue;
                    }

                    int ret = sapModel.PropFrame.SetRectangle(input.SectionName, input.MaterialName, input.H, input.B);
                    if (ret == 0) ctx.IncrementRan(); else { failCount++; ctx.IncrementSkipped(); }
                }
            });

            string msg = $"Created: {progress.RanCount}, Skipped: {progress.SkippedCount}, Failed: {failCount}";
            if (progress.WasCancelled) msg += " (Cancelled)";
            return OperationResult.Success(msg);
        }

        public OperationResult AddConcreteCircleSections(IReadOnlyList<CSISapModelConcreteCircleSectionInput> inputs)
        {
            ETABSv1.cSapModel sapModel = _currentConnection?.SapModel as ETABSv1.cSapModel;
            if (sapModel == null) return OperationResult.Failure("No active ETABS model found.");

            int unitRet = sapModel.SetPresentUnits(ETABSv1.eUnits.N_mm_C);
            if (unitRet != 0) return OperationResult.Failure("Failed to set ETABS present units to N-mm-C.");

            int failCount = 0;

            var progress = BatchProgressWindow.RunWithProgress(inputs.Count, "Creating Concrete Circle Sections...", (ctx) =>
            {
                foreach (var input in inputs)
                {
                    if (ctx.IsCancellationRequested) break;

                    if (FrameSectionExists(sapModel, input.SectionName))
                    {
                        ctx.IncrementSkipped();
                        continue;
                    }

                    int ret = sapModel.PropFrame.SetCircle(input.SectionName, input.MaterialName, input.D);
                    if (ret == 0) ctx.IncrementRan(); else { failCount++; ctx.IncrementSkipped(); }
                }
            });

            string msg = $"Created: {progress.RanCount}, Skipped: {progress.SkippedCount}, Failed: {failCount}";
            if (progress.WasCancelled) msg += " (Cancelled)";
            return OperationResult.Success(msg);
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
                ETABSv1.cSapModel sapModel = (ETABSv1.cSapModel)connectionResult.Data.SapModel;

                int unitRet = sapModel.SetPresentUnits(ETABSv1.eUnits.kN_m_C);
                if (unitRet != 0)
                {
                    return OperationResult.Failure("Failed to set ETABS present units to kN-m-C.");
                }

                var framesResult = ReadSelectedFrameGeometries(sapModel);
                if (!framesResult.IsSuccess)
                {
                    return OperationResult.Failure(framesResult.Message);
                }

                var frameGeometries = framesResult.Data;
                if (frameGeometries == null || frameGeometries.Count == 0)
                {
                    return OperationResult.Failure("No frame objects are currently selected in ETABS.");
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
                            if (ctx.IsCancellationRequested)
                            {
                                break;
                            }

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
                return OperationResult.Failure($"ETABS COM error while creating shell areas from selected frames: {ex.Message}");
            }
            catch (Exception ex)
            {
                return OperationResult.Failure($"Failed to create shell areas from selected frames: {ex.Message}");
            }
        }

        private OperationResult<IReadOnlyList<ShellFrameGeometry>> ReadSelectedFrameGeometries(ETABSv1.cSapModel sapModel)
        {
            int numberItems = 0;
            int[] objectTypes = null;
            string[] objectNames = null;
            int selectedResult = sapModel.SelectObj.GetSelected(ref numberItems, ref objectTypes, ref objectNames);

            if (selectedResult != 0)
            {
                return OperationResult<IReadOnlyList<ShellFrameGeometry>>.Failure("Failed to read selected objects from ETABS.");
            }

            if (numberItems <= 0 || objectTypes == null || objectNames == null)
            {
                return OperationResult<IReadOnlyList<ShellFrameGeometry>>.Failure("No frame objects are currently selected in ETABS.");
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
                if (objectTypes[i] != CSISapModelObjectTypeIds.Frame ||
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
                return OperationResult<IReadOnlyList<ShellFrameGeometry>>.Failure("No valid frame geometry could be read from the current ETABS selection.");
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
            if (nodeCount == 4)
            {
                return 0;
            }

            if (nodeCount == 3)
            {
                return 1;
            }

            return 2;
        }

        private static int CreateAreaFromLoop(
            ETABSv1.cSapModel sapModel,
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

                return SplitQuadAndCreateTwoTriangles(
                    sapModel,
                    loopPts,
                    pointCoords,
                    propName,
                    acceptedFaces,
                    tolerances);
            }

            return 0;
        }

        private static bool AddAreaByNodeCoordinates(
            ETABSv1.cSapModel sapModel,
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
            ETABSv1.cSapModel sapModel,
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

        private class ShellFaceCandidate
        {
            public string[] OrderedLoop { get; set; }
            public double Area { get; set; }
        }

        private bool FrameSectionExists(ETABSv1.cSapModel sapModel, string sectionName)
        {
            if (string.IsNullOrWhiteSpace(sectionName))
            {
                return false;
            }

            ETABSv1.eFramePropType propType = ETABSv1.eFramePropType.I;
            int ret = sapModel.PropFrame.GetTypeOAPI(sectionName, ref propType);
            return ret == 0;
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
