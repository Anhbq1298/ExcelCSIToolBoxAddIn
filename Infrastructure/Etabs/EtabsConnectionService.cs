using System;
using System.Collections.Generic;
using System.IO;
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

        public OperationResult<IReadOnlyList<EtabsPointData>> GetSelectedPointsFromActiveModel()
        {
            var connectionResult = GetCurrentConnection();
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
