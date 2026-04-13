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
        private const int PointObjectType = 1;

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
                string modelPath = GetModelFileNameSafely(sapModel);
                string modelName = string.IsNullOrWhiteSpace(modelPath)
                    ? "Unknown model"
                    : Path.GetFileName(modelPath);

                _currentConnection = new EtabsConnectionInfo
                {
                    IsConnected = true,
                    ModelFileName = modelName,
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
                        false,
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

                    if (objectTypes[i] != PointObjectType || string.IsNullOrWhiteSpace(objectNames[i]))
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

        private static string GetModelFileNameSafely(ETABSv1.cSapModel sapModel)
        {
            if (sapModel == null)
            {
                return null;
            }

            try
            {
                return sapModel.GetModelFilename(true);
            }
            catch
            {
                try
                {
                    return sapModel.GetModelFilename();
                }
                catch
                {
                    return null;
                }
            }
        }
    }
}
