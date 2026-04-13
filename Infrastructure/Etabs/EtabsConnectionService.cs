using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
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
                dynamic etabsObject = _currentConnection.EtabsObject;
                int result = etabsObject.ApplicationExit(false);

                _currentConnection = null;

                if (result != 0)
                {
                    return OperationResult.Failure("Failed to close the attached ETABS instance.");
                }

                return OperationResult.Success("Successfully closed the attached ETABS instance.");
            }
            catch
            {
                _currentConnection = null;
                return OperationResult.Failure("Failed to close the attached ETABS instance.");
            }
        }

        public OperationResult SelectPointsByUniqueNames(IReadOnlyList<string> uniqueNames)
        {
            if (uniqueNames == null || uniqueNames.Count == 0)
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
                dynamic sapModel = connectionResult.Data.SapModel;
                sapModel.SelectObj.ClearSelection();

                var unresolved = new List<string>();
                var selectedCount = 0;

                foreach (var uniqueName in uniqueNames.Where(x => !string.IsNullOrWhiteSpace(x)).Distinct(StringComparer.OrdinalIgnoreCase))
                {
                    int result = sapModel.PointObj.SetSelected(uniqueName, true);
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

                return OperationResult.Success(message);
            }
            catch
            {
                return OperationResult.Failure("Failed to select ETABS points by UniqueName.");
            }
        }

        public OperationResult SelectPointsByLabels(IReadOnlyList<string> labels)
        {
            if (labels == null || labels.Count == 0)
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
                dynamic sapModel = connectionResult.Data.SapModel;
                sapModel.SelectObj.ClearSelection();

                var pointLookupResult = BuildPointLookupByLabel(sapModel);
                if (!pointLookupResult.IsSuccess)
                {
                    return OperationResult.Failure(pointLookupResult.Message);
                }

                var labelsByUniqueName = pointLookupResult.Data;
                var unresolvedLabels = new List<string>();
                var selectedCount = 0;

                foreach (var label in labels.Where(x => !string.IsNullOrWhiteSpace(x)).Distinct(StringComparer.OrdinalIgnoreCase))
                {
                    string pointUniqueName;
                    if (!labelsByUniqueName.TryGetValue(label, out pointUniqueName))
                    {
                        unresolvedLabels.Add(label);
                        continue;
                    }

                    int result = sapModel.PointObj.SetSelected(pointUniqueName, true);
                    if (result == 0)
                    {
                        selectedCount++;
                    }
                    else
                    {
                        unresolvedLabels.Add(label);
                    }
                }

                var message = $"Selected {selectedCount} point(s) by Label.";
                if (unresolvedLabels.Count > 0)
                {
                    message += $" Not found: {string.Join(", ", unresolvedLabels)}.";
                }

                return OperationResult.Success(message);
            }
            catch
            {
                return OperationResult.Failure("Failed to select ETABS points by Label.");
            }
        }

        public OperationResult<EtabsAddPointsResult> AddPointsCartesian(IReadOnlyList<EtabsPointCartesianInput> points)
        {
            if (points == null || points.Count == 0)
            {
                return OperationResult<EtabsAddPointsResult>.Failure("No valid rows were found. Please verify X, Y, Z are numeric.");
            }

            var connectionResult = EnsureConnection();
            if (!connectionResult.IsSuccess || connectionResult.Data?.SapModel == null)
            {
                return OperationResult<EtabsAddPointsResult>.Failure(connectionResult.Message);
            }

            try
            {
                dynamic sapModel = connectionResult.Data.SapModel;
                var failedRows = new List<int>();
                var successCount = 0;

                foreach (var point in points)
                {
                    string pointName = string.IsNullOrWhiteSpace(point.Name) ? string.Empty : point.Name;
                    int result = sapModel.PointObj.AddCartesian(point.X, point.Y, point.Z, ref pointName, pointName, "Global");

                    if (result == 0)
                    {
                        successCount++;
                    }
                    else
                    {
                        failedRows.Add(point.ExcelRowNumber);
                    }
                }

                var data = new EtabsAddPointsResult
                {
                    AddedCount = successCount,
                    FailedRows = failedRows
                };

                return OperationResult<EtabsAddPointsResult>.Success(data);
            }
            catch
            {
                return OperationResult<EtabsAddPointsResult>.Failure("Failed to add points to ETABS from the selected Excel range.");
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

                    if (pointResult == 0)
                    {
                        points.Add(new EtabsPointData
                        {
                            PointUniqueName = objectNames[i],
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

        private static OperationResult<Dictionary<string, string>> BuildPointLookupByLabel(dynamic sapModel)
        {
            // Label-based API calls may require story context, so we map labels by reading all point names
            // and their labels from the active model to stay consistent and predictable.
            int pointCount = 0;
            string[] pointNames = null;
            int namesResult = sapModel.PointObj.GetNameList(ref pointCount, ref pointNames);

            if (namesResult != 0)
            {
                return OperationResult<Dictionary<string, string>>.Failure("Failed to read ETABS point names for label selection.");
            }

            var labelsByUniqueName = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            for (int i = 0; i < pointCount; i++)
            {
                if (pointNames == null || i >= pointNames.Length || string.IsNullOrWhiteSpace(pointNames[i]))
                {
                    continue;
                }

                string label = string.Empty;
                string story = string.Empty;
                int labelResult = sapModel.PointObj.GetLabelFromName(pointNames[i], ref label, ref story);
                if (labelResult == 0 && !string.IsNullOrWhiteSpace(label) && !labelsByUniqueName.ContainsKey(label))
                {
                    labelsByUniqueName[label] = pointNames[i];
                }
            }

            return OperationResult<Dictionary<string, string>>.Success(labelsByUniqueName);
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
