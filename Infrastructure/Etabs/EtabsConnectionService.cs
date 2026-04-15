using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using ExcelCSIToolBoxAddIn.Common.Results;

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
                        string forceUnit;
                        switch (GetEnumKeyName(forceUnits).ToUpperInvariant())
                        {
                            case "KN":
                                forceUnit = "kN";
                                break;
                            case "KIP":
                                forceUnit = "kip";
                                break;
                            case "LB":
                                forceUnit = "lb";
                                break;
                            case "N":
                                forceUnit = "N";
                                break;
                            case "KGF":
                                forceUnit = "kgf";
                                break;
                            case "TONF":
                                forceUnit = "tonf";
                                break;
                            default:
                                forceUnit = GetEnumKeyName(forceUnits);
                                break;
                        }

                        string lengthUnit;
                        switch (GetEnumKeyName(lengthUnits).ToUpperInvariant())
                        {
                            case "M":
                                lengthUnit = "m";
                                break;
                            case "MM":
                                lengthUnit = "mm";
                                break;
                            case "CM":
                                lengthUnit = "cm";
                                break;
                            case "FT":
                                lengthUnit = "ft";
                                break;
                            case "INCH":
                                lengthUnit = "inch";
                                break;
                            case "MICRON":
                                lengthUnit = "micron";
                                break;
                            default:
                                lengthUnit = GetEnumKeyName(lengthUnits);
                                break;
                        }

                        string temperatureUnit;
                        switch (GetEnumKeyName(temperatureUnits).ToUpperInvariant())
                        {
                            case "C":
                                temperatureUnit = "C";
                                break;
                            case "F":
                                temperatureUnit = "F";
                                break;
                            default:
                                temperatureUnit = GetEnumKeyName(temperatureUnits);
                                break;
                        }

                        modelCurrentUnit = $"{forceUnit}-{lengthUnit}-{temperatureUnit}";
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
                    return OperationResult.Failure("Failed to clear ETABS selection before selecting points by UniqueName.");
                }

                var unresolved = new List<string>();
                var selectedCount = 0;

                foreach (var uniqueName in orderedUniqueNames)
                {
                    int result = sapModel.PointObj.SetSelected(uniqueName, true, ETABSv1.eItemType.Objects);
                    if (result == 0)
                    {
                        selectedCount++;
                    }
                    else
                    {
                        unresolved.Add(uniqueName);
                    }
                }

                var message = $"Selected {selectedCount} point(s) by UniqueName.";
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
                return OperationResult.Failure("Failed to select ETABS points by UniqueName.");
            }
        }

        public OperationResult SelectFramesByUniqueNames(IReadOnlyList<string> uniqueNames)
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
                    return OperationResult.Failure("Failed to clear ETABS selection before selecting frames by UniqueName.");
                }

                var unresolved = new List<string>();
                var selectedCount = 0;

                foreach (var uniqueName in orderedUniqueNames)
                {
                    int result = sapModel.FrameObj.SetSelected(uniqueName, true, ETABSv1.eItemType.Objects);
                    if (result == 0)
                    {
                        selectedCount++;
                    }
                    else
                    {
                        unresolved.Add(uniqueName);
                    }
                }

                var message = $"Selected {selectedCount} frame(s) by UniqueName.";
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
                return OperationResult.Failure("Failed to select ETABS frames by UniqueName.");
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
                // no grouping and no de-duplication.
                // IMPORTANT: PointObj.AddCartesian is called with MergeOff = true so ETABS
                // does NOT merge points even when multiple rows share identical coordinates.
                foreach (var pointInput in pointInputs)
                {
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
                        continue;
                    }

                    successCount++;

                    if (!string.IsNullOrWhiteSpace(requestedUniqueName) &&
                        !string.Equals(assignedName, requestedUniqueName, StringComparison.OrdinalIgnoreCase))
                    {
                        failedRowMessages.Add(
                            $"Row {pointInput.ExcelRowNumber}: Point was created, but ETABS assigned UniqueName '{assignedName}' instead of requested '{requestedUniqueName}'.");
                    }
                }

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

                foreach (var frameInput in frameInputs)
                {
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
                        continue;
                    }

                    successCount++;
                }

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

                foreach (var frameInput in frameInputs)
                {
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
                        continue;
                    }

                    successCount++;

                    if (!string.IsNullOrWhiteSpace(userName) &&
                        !string.Equals(createdName, userName, StringComparison.OrdinalIgnoreCase))
                    {
                        failedRowMessages.Add(
                            $"Row {frameInput.ExcelRowNumber}: Frame was created, but ETABS assigned UniqueName '{createdName}' instead of requested '{userName}'.");
                    }
                }

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

        private static OperationResult RefreshView(ETABSv1.cSapModel sapModel)
        {
            int refreshResult = sapModel.View.RefreshView(0, false);
            if (refreshResult != 0)
            {
                return OperationResult.Failure($"ETABS model changed successfully, but View.RefreshView failed (return code {refreshResult}).");
            }

            return OperationResult.Success();
        }

        private static string GetEnumKeyName<TEnum>(TEnum enumValue) where TEnum : struct
        {
            var enumType = typeof(TEnum);
            var enumName = Enum.GetName(enumType, enumValue);
            if (!string.IsNullOrWhiteSpace(enumName))
            {
                return enumName;
            }

            if (!enumType.IsEnum)
            {
                return "?";
            }

            return Convert.ToInt32(enumValue).ToString();
        }

        private ETABSv1.cSapModel GetActiveSapModel(EtabsConnectionInfo connectionInfo)
        {
            if (connectionInfo == null)
            {
                return null;
            }

            var etabsApplication = connectionInfo.EtabsObject as ETABSv1.cOAPI;
            if (etabsApplication != null)
            {
                try
                {
                    var activeSapModel = etabsApplication.SapModel;
                    if (activeSapModel != null)
                    {
                        connectionInfo.SapModel = activeSapModel;
                        return activeSapModel;
                    }
                }
                catch
                {
                    // Fallback to existing SapModel reference below.
                }
            }

            return connectionInfo.SapModel as ETABSv1.cSapModel;
        }
    }
}
