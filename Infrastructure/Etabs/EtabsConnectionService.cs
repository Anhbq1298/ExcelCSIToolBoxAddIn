using System;
using System.Collections.Generic;
using System.IO;
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
            var sapModelResult = EnsureEtabsSapModel();
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
                (sapModel, name) => sapModel.PointObj.SetSelected(name, true, ETABSv1.eItemType.Objects),
                RefreshView);
        }

        public OperationResult SelectFramesByUniqueNames(IReadOnlyList<string> uniqueNames)
        {
            var sapModelResult = EnsureEtabsSapModel();
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
                (sapModel, name) => sapModel.FrameObj.SetSelected(name, true, ETABSv1.eItemType.Objects),
                RefreshView);
        }

        public OperationResult<CSISapModelAddPointsResult> AddPointsByCartesian(IReadOnlyList<CSISapModelPointCartesianInput> pointInputs)
        {
            var sapModelResult = EnsureEtabsSapModel();
            if (!sapModelResult.IsSuccess)
            {
                return OperationResult<CSISapModelAddPointsResult>.Failure(sapModelResult.Message);
            }

            return CSISapModelOperationRunner.AddPointsByCartesian(
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
        }

        public OperationResult<CSISapModelAddFramesResult> AddFramesByCoordinates(IReadOnlyList<CSISapModelFrameByCoordInput> frameInputs)
        {
            var sapModelResult = EnsureEtabsSapModel();
            if (!sapModelResult.IsSuccess)
            {
                return OperationResult<CSISapModelAddFramesResult>.Failure(sapModelResult.Message);
            }

            return CSISapModelOperationRunner.AddFramesByCoordinates(
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
        }

        public OperationResult<CSISapModelAddFramesResult> AddFramesByPoint(IReadOnlyList<CSISapModelFrameByPointInput> frameInputs)
        {
            var sapModelResult = EnsureEtabsSapModel();
            if (!sapModelResult.IsSuccess)
            {
                return OperationResult<CSISapModelAddFramesResult>.Failure(sapModelResult.Message);
            }

            return CSISapModelOperationRunner.AddFramesByPoint(
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

        private OperationResult<ETABSv1.cSapModel> EnsureEtabsSapModel()
        {
            var connectionResult = EnsureConnection();
            if (!connectionResult.IsSuccess || connectionResult.Data?.SapModel == null)
            {
                return OperationResult<ETABSv1.cSapModel>.Failure(connectionResult.Message);
            }

            var sapModel = connectionResult.Data.SapModel as ETABSv1.cSapModel;
            if (sapModel == null)
            {
                return OperationResult<ETABSv1.cSapModel>.Failure("The attached ETABS SapModel is invalid. Please reattach and try again.");
            }

            return OperationResult<ETABSv1.cSapModel>.Success(sapModel);
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
            var sapModelResult = EnsureEtabsSapModel();
            if (!sapModelResult.IsSuccess)
            {
                return OperationResult.Failure(sapModelResult.Message);
            }

            return CSISapModelShellAreaService.CreateShellAreasFromSelectedFrames(
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
