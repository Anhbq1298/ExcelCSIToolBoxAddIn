using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using ExcelCSIToolBoxAddIn.Common.Results;
using ExcelCSIToolBoxAddIn.Core.Geometry;
using ExcelCSIToolBoxAddIn.UI.Views;

namespace ExcelCSIToolBoxAddIn.Infrastructure.Etabs
{
    /// <summary>
    /// Infrastructure service that safely attaches to a running ETABS instance.
    /// Stores the latest attached instance so ETABS commands can reuse the same SapModel.
    /// </summary>
    public class EtabsConnectionService : IEtabsConnectionService
    {
        private const string EtabsComProgId = "CSI.ETABS.API.ETABSObject";

        private EtabsConnectionInfo _currentConnection;

        public OperationResult<EtabsConnectionInfo> TryAttachToRunningInstance()
        {
            ETABSv1.cHelper myHelper = new ETABSv1.Helper();
            ETABSv1.cOAPI myETABSObject = null;

            try
            {
                myETABSObject = myHelper.GetObject(EtabsComProgId);

                if (myETABSObject == null)
                {
                    _currentConnection = null;
                    return OperationResult<EtabsConnectionInfo>.Failure("ETABS is not running.");
                }

                ETABSv1.cSapModel sapModel = myETABSObject.SapModel;

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

                _currentConnection = new EtabsConnectionInfo
                {
                    IsConnected = true,
                    ModelPath = modelPath,
                    ModelFileName = modelName,
                    ModelCurrentUnit = modelCurrentUnit,
                    EtabsObject = myETABSObject,
                    SapModel = sapModel
                };

                return OperationResult<EtabsConnectionInfo>.Success(_currentConnection);
            }
            catch
            {
                _currentConnection = null;
                return OperationResult<EtabsConnectionInfo>.Failure("ETABS is not running.");
            }
        }

        public OperationResult<EtabsConnectionInfo> GetCurrentConnection()
        {
            if (_currentConnection?.SapModel == null)
            {
                return OperationResult<EtabsConnectionInfo>.Failure("No ETABS model is currently connected. Please click 'Attach to Running ETABS'.");
            }

            return OperationResult<EtabsConnectionInfo>.Success(_currentConnection);
        }

        public OperationResult CloseCurrentEtabsInstance()
        {
            if (_currentConnection?.EtabsObject == null)
            {
                return OperationResult.Failure("No running ETABS instance is currently attached.");
            }

            try
            {
                var etabsApplication = _currentConnection.EtabsObject as ETABSv1.cOAPI;
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
            ReleaseComReference(_currentConnection.EtabsObject);

            _currentConnection.SapModel = null;
            _currentConnection.EtabsObject = null;
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
            return SelectObjectsByUniqueNames(uniqueNames, "point",
                (sapModel, name) => sapModel.PointObj.SetSelected(name, true, ETABSv1.eItemType.Objects));
        }

        public OperationResult SelectFramesByUniqueNames(IReadOnlyList<string> uniqueNames)
        {
            return SelectObjectsByUniqueNames(uniqueNames, "frame",
                (sapModel, name) => sapModel.FrameObj.SetSelected(name, true, ETABSv1.eItemType.Objects));
        }

        private OperationResult SelectObjectsByUniqueNames(
            IReadOnlyList<string> uniqueNames, string objectTypeName,
            Func<ETABSv1.cSapModel, string, int> setSelected)
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
                ETABSv1.cSapModel sapModel = (ETABSv1.cSapModel)connectionResult.Data.SapModel;

                int clearSelectionResult = sapModel.SelectObj.ClearSelection();
                if (clearSelectionResult != 0)
                {
                    return OperationResult.Failure($"Failed to clear ETABS selection before selecting {objectTypeName}s by UniqueName.");
                }

                var unresolved = new List<string>();
                var selectedCount = 0;

                var progress = BatchProgressWindow.RunWithProgress(orderedUniqueNames.Count, $"Selecting {objectTypeName}s...", (ctx) =>
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
                return OperationResult.Failure($"Failed to select ETABS {objectTypeName}s by UniqueName.");
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

        public OperationResult<EtabsAddPointsResult> AddPointsByCartesian(IReadOnlyList<EtabsPointCartesianInput> pointInputs)
        {
            if (pointInputs == null || pointInputs.Count == 0)
            {
                return OperationResult<EtabsAddPointsResult>.Failure("No valid rows were found in the selected range.");
            }

            var connectionResult = EnsureConnection();
            if (!connectionResult.IsSuccess || connectionResult.Data?.SapModel == null)
            {
                return OperationResult<EtabsAddPointsResult>.Failure(connectionResult.Message);
            }

            try
            {
                ETABSv1.cSapModel sapModel = (ETABSv1.cSapModel)connectionResult.Data.SapModel;
                var failedRowMessages = new List<string>();
                var successCount = 0;

                // Process each row exactly as provided by Excel parsing:
                var progress = BatchProgressWindow.RunWithProgress(pointInputs.Count, "Adding Points to Model...", (ctx) =>
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
                            failedRowMessages.Add(
                                $"Row {pointInput.ExcelRowNumber}: ETABS API call PointObj.AddCartesian failed (return code {addResult}).");
                            ctx.IncrementSkipped();
                            continue;
                        }

                        successCount++;
                        ctx.IncrementRan();

                        if (!string.IsNullOrWhiteSpace(requestedUniqueName) &&
                            !string.Equals(assignedName, requestedUniqueName, StringComparison.OrdinalIgnoreCase))
                        {
                            failedRowMessages.Add(
                                $"Row {pointInput.ExcelRowNumber}: Point was created, but ETABS assigned UniqueName '{assignedName}' instead of requested '{requestedUniqueName}'.");
                        }
                    }
                });

                var data = new EtabsAddPointsResult
                {
                    AddedCount = successCount,
                    FailedRowMessages = failedRowMessages
                };

                if (successCount > 0)
                {
                    var refreshResult = RefreshView(sapModel);
                    if (!refreshResult.IsSuccess)
                    {
                        return OperationResult<EtabsAddPointsResult>.Failure(refreshResult.Message);
                    }
                }

                return OperationResult<EtabsAddPointsResult>.Success(data);
            }
            catch (COMException ex)
            {
                return OperationResult<EtabsAddPointsResult>.Failure($"ETABS COM error while adding points by Cartesian coordinates: {ex.Message}");
            }
            catch (Exception ex)
            {
                return OperationResult<EtabsAddPointsResult>.Failure(
                    $"ETABS add-by-Cartesian failed unexpectedly: {ex.Message}");
            }
        }

        public OperationResult<EtabsAddFramesResult> AddFramesByCoordinates(IReadOnlyList<EtabsFrameByCoordInput> frameInputs)
        {
            if (frameInputs == null || frameInputs.Count == 0)
            {
                return OperationResult<EtabsAddFramesResult>.Failure("No valid rows were found in the selected range.");
            }

            var connectionResult = EnsureConnection();
            if (!connectionResult.IsSuccess || connectionResult.Data?.SapModel == null)
            {
                return OperationResult<EtabsAddFramesResult>.Failure(connectionResult.Message);
            }

            try
            {
                ETABSv1.cSapModel sapModel = (ETABSv1.cSapModel)connectionResult.Data.SapModel;
                var failedRowMessages = new List<string>();
                var successCount = 0;

                var progress = BatchProgressWindow.RunWithProgress(frameInputs.Count, "Adding Frames (by Coordinates)...", (ctx) =>
                {
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
                            failedRowMessages.Add(
                                $"Row {frameInput.ExcelRowNumber}: ETABS API call FrameObj.AddByCoord failed (return code {addResult}).");
                            ctx.IncrementSkipped();
                            continue;
                        }

                        successCount++;
                        ctx.IncrementRan();
                    }
                });

                var data = new EtabsAddFramesResult
                {
                    AddedCount = successCount,
                    FailedRowMessages = failedRowMessages
                };

                if (successCount > 0)
                {
                    var refreshResult = RefreshView(sapModel);
                    if (!refreshResult.IsSuccess)
                    {
                        return OperationResult<EtabsAddFramesResult>.Failure(refreshResult.Message);
                    }
                }

                return OperationResult<EtabsAddFramesResult>.Success(data);
            }
            catch (COMException ex)
            {
                return OperationResult<EtabsAddFramesResult>.Failure($"ETABS COM error while adding frames by coordinates: {ex.Message}");
            }
            catch (Exception ex)
            {
                return OperationResult<EtabsAddFramesResult>.Failure(
                    $"ETABS add-by-coordinates failed unexpectedly: {ex.Message}");
            }
        }

        public OperationResult<EtabsAddFramesResult> AddFramesByPoint(IReadOnlyList<EtabsFrameByPointInput> frameInputs)
        {
            if (frameInputs == null || frameInputs.Count == 0)
            {
                return OperationResult<EtabsAddFramesResult>.Failure("No valid rows were found in the selected range.");
            }

            var connectionResult = EnsureConnection();
            if (!connectionResult.IsSuccess || connectionResult.Data?.SapModel == null)
            {
                return OperationResult<EtabsAddFramesResult>.Failure(connectionResult.Message);
            }

            try
            {
                ETABSv1.cSapModel sapModel = (ETABSv1.cSapModel)connectionResult.Data.SapModel;
                var failedRowMessages = new List<string>();
                var successCount = 0;

                var progress = BatchProgressWindow.RunWithProgress(frameInputs.Count, "Adding Frames (by Point Names)...", (ctx) =>
                {
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
                            failedRowMessages.Add(
                                $"Row {frameInput.ExcelRowNumber}: ETABS API call FrameObj.AddByPoint failed (return code {addResult}).");
                            ctx.IncrementSkipped();
                            continue;
                        }

                        successCount++;
                        ctx.IncrementRan();

                        if (!string.IsNullOrWhiteSpace(userName) &&
                            !string.Equals(createdName, userName, StringComparison.OrdinalIgnoreCase))
                        {
                            failedRowMessages.Add(
                                $"Row {frameInput.ExcelRowNumber}: Frame was created, but ETABS assigned UniqueName '{createdName}' instead of requested '{userName}'.");
                        }
                    }
                });

                var data = new EtabsAddFramesResult
                {
                    AddedCount = successCount,
                    FailedRowMessages = failedRowMessages
                };

                if (successCount > 0)
                {
                    var refreshResult = RefreshView(sapModel);
                    if (!refreshResult.IsSuccess)
                    {
                        return OperationResult<EtabsAddFramesResult>.Failure(refreshResult.Message);
                    }
                }

                return OperationResult<EtabsAddFramesResult>.Success(data);
            }
            catch (COMException ex)
            {
                return OperationResult<EtabsAddFramesResult>.Failure($"ETABS COM error while adding frames by point names: {ex.Message}");
            }
            catch (Exception ex)
            {
                return OperationResult<EtabsAddFramesResult>.Failure(
                    $"ETABS add-by-point failed unexpectedly: {ex.Message}");
            }
        }

        public OperationResult<IReadOnlyList<EtabsPointData>> GetSelectedPointsFromActiveModel()
        {
            var connectionResult = EnsureConnection();
            if (!connectionResult.IsSuccess || connectionResult.Data?.SapModel == null)
            {
                return OperationResult<IReadOnlyList<EtabsPointData>>.Failure(connectionResult.Message);
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
                    return OperationResult<IReadOnlyList<EtabsPointData>>.Failure("Failed to read selected objects from ETABS.");
                }

                var points = new List<EtabsPointData>();

                for (int i = 0; i < numberItems; i++)
                {
                    if (objectTypes == null || objectNames == null || i >= objectTypes.Length || i >= objectNames.Length)
                    {
                        continue;
                    }

                    if (objectTypes[i] != EtabsObjectTypeIds.Point || string.IsNullOrWhiteSpace(objectNames[i]))
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
                        points.Add(new EtabsPointData
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
                    return OperationResult<IReadOnlyList<EtabsPointData>>.Failure("No point objects are selected in ETABS.");
                }

                return OperationResult<IReadOnlyList<EtabsPointData>>.Success(points);
            }
            catch
            {
                return OperationResult<IReadOnlyList<EtabsPointData>>.Failure("Unable to read selected ETABS points.");
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
                    if (objectTypes[i] != EtabsObjectTypeIds.Frame || string.IsNullOrWhiteSpace(frameUniqueName))
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


        private OperationResult<EtabsConnectionInfo> EnsureConnection()
        {
            var connectionResult = GetCurrentConnection();
            if (!connectionResult.IsSuccess || connectionResult.Data?.SapModel == null)
            {
                connectionResult = TryAttachToRunningInstance();
            }

            return connectionResult;
        }

        public OperationResult AddSteelISections(IReadOnlyList<EtabsSteelISectionInput> inputs)
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

        public OperationResult AddSteelChannelSections(IReadOnlyList<EtabsSteelChannelSectionInput> inputs)
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

        public OperationResult AddSteelAngleSections(IReadOnlyList<EtabsSteelAngleSectionInput> inputs)
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

        public OperationResult AddSteelPipeSections(IReadOnlyList<EtabsSteelPipeSectionInput> inputs)
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

        public OperationResult AddSteelTubeSections(IReadOnlyList<EtabsSteelTubeSectionInput> inputs)
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

        public OperationResult AddConcreteRectangleSections(IReadOnlyList<EtabsConcreteRectangleSectionInput> inputs)
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

        public OperationResult AddConcreteCircleSections(IReadOnlyList<EtabsConcreteCircleSectionInput> inputs)
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
                if (objectTypes[i] != EtabsObjectTypeIds.Frame ||
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
                    OrderedLoop = orderedLoop
                });
            }

            return candidates
                .OrderBy(candidate => GetShellLoopPriority(candidate.OrderedLoop.Length))
                .ThenBy(candidate => candidate.OrderedLoop.Length)
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
