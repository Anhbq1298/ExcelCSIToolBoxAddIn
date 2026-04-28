using System;
using System.IO;
using System.Runtime.InteropServices;
using ExcelCSIToolBoxAddIn.Adapters;
using ExcelCSIToolBoxAddIn.Common.Results;
using ExcelCSIToolBoxAddIn.Infrastructure.Etabs;
using ExcelCSIToolBoxAddIn.Infrastructure.Sap2000;

namespace ExcelCSIToolBoxAddIn.Infrastructure.CSISapModel
{
    internal class CSISapModelConnectionAdapter<TSapModel, TCsiObject> : ICSISapModelConnectionAdapter<TSapModel>
        where TSapModel : class
        where TCsiObject : class
    {
        private readonly ICsiModelAdapter _modelAdapter;
        private readonly Func<TSapModel, string> _getModelPath;
        private readonly Func<TSapModel, string> _getModelCurrentUnit;
        private readonly Func<TCsiObject, int> _closeApplication;
        private CSISapModelConnectionInfoDTO _currentConnection;

        internal CSISapModelConnectionAdapter(
            ICsiModelAdapter modelAdapter,
            string productName,
            Func<TSapModel, string> getModelPath,
            Func<TSapModel, string> getModelCurrentUnit,
            Func<TCsiObject, int> closeApplication)
        {
            _modelAdapter = modelAdapter ?? throw new ArgumentNullException(nameof(modelAdapter));
            ProductName = productName ?? throw new ArgumentNullException(nameof(productName));
            _getModelPath = getModelPath ?? throw new ArgumentNullException(nameof(getModelPath));
            _getModelCurrentUnit = getModelCurrentUnit ?? throw new ArgumentNullException(nameof(getModelCurrentUnit));
            _closeApplication = closeApplication ?? throw new ArgumentNullException(nameof(closeApplication));
        }

        public string ProductName { get; }

        public OperationResult<CSISapModelConnectionInfoDTO> TryAttachToRunningInstance()
        {
            var attachResult = _modelAdapter.AttachToRunningInstance();
            if (!attachResult.IsSuccess)
            {
                _currentConnection = null;
                return OperationResult<CSISapModelConnectionInfo>.Failure(attachResult.Message);
            }

            var csiObject = attachResult.ApplicationObject as TCsiObject;
            var sapModel = attachResult.SapModel as TSapModel;
            if (csiObject == null || sapModel == null)
            {
                _currentConnection = null;
                return OperationResult<CSISapModelConnectionInfo>.Failure($"The attached {ProductName} instance is invalid. Please reattach and try again.");
            }

            try
            {
                var modelPath = _getModelPath(sapModel);
                _currentConnection = new CSISapModelConnectionInfoDTO
                {
                    IsConnected = true,
                    ModelPath = modelPath,
                    ModelFileName = string.IsNullOrWhiteSpace(modelPath) ? "Unsaved Model" : Path.GetFileName(modelPath),
                    ModelCurrentUnit = _getModelCurrentUnit(sapModel),
                    CsiObject = csiObject,
                    SapModel = sapModel
                };

                return OperationResult<CSISapModelConnectionInfoDTO>.Success(_currentConnection);
            }
            catch
            {
                _currentConnection = null;
                return OperationResult<CSISapModelConnectionInfoDTO>.Failure($"Failed to attach to the running {ProductName} instance.");
            }
        }

        public OperationResult<CSISapModelConnectionInfoDTO> GetCurrentConnection()
        {
            if (_currentConnection?.SapModel == null)
            {
                return OperationResult<CSISapModelConnectionInfoDTO>.Failure($"No {ProductName} model is currently connected. Please attach to a running {ProductName} instance.");
            }

            return OperationResult<CSISapModelConnectionInfo>.Success(_currentConnection);
        }

        public OperationResult<TSapModel> EnsureSapModel()
        {
            var connectionResult = GetCurrentConnection();
            if (!connectionResult.IsSuccess || connectionResult.Data?.SapModel == null)
            {
                connectionResult = TryAttachToRunningInstance();
            }

            if (!connectionResult.IsSuccess || connectionResult.Data?.SapModel == null)
            {
                return OperationResult<TSapModel>.Failure(connectionResult.Message);
            }

            var sapModel = connectionResult.Data.SapModel as TSapModel;
            if (sapModel == null)
            {
                return OperationResult<TSapModel>.Failure($"The attached {ProductName} SapModel is invalid. Please reattach and try again.");
            }

            return OperationResult<TSapModel>.Success(sapModel);
        }

        public OperationResult CloseCurrentInstance()
        {
            if (_currentConnection?.CsiObject == null)
            {
                return OperationResult.Failure($"No running {ProductName} instance is currently attached.");
            }

            try
            {
                var csiObject = _currentConnection.CsiObject as TCsiObject;
                if (csiObject == null)
                {
                    return OperationResult.Failure($"The attached {ProductName} instance is invalid. Please reattach and try again.");
                }

                int result = _closeApplication(csiObject);
                if (result != 0)
                {
                    return OperationResult.Failure($"{ProductName} failed to close the attached instance (ApplicationExit returned {result}).");
                }

                ResetCurrentConnection();
                return OperationResult.Success($"Successfully closed the attached {ProductName} instance.");
            }
            catch (COMException ex)
            {
                return OperationResult.Failure($"{ProductName} COM error while closing attached instance: {ex.Message}");
            }
            catch
            {
                return OperationResult.Failure($"Failed to close the attached {ProductName} instance.");
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
            }
        }
    }

    internal static class CSISapModelConnectionAdapterFactory
    {
        internal static ICSISapModelConnectionAdapter<ETABSv1.cSapModel> CreateEtabs()
        {
            return CreateEtabs(new EtabsModelAdapter());
        }

        internal static ICSISapModelConnectionAdapter<ETABSv1.cSapModel> CreateEtabs(ICsiModelAdapter modelAdapter)
        {
            var adapter = new CSISapModelConnectionAdapter<ETABSv1.cSapModel, ETABSv1.cOAPI>(
                modelAdapter,
                "ETABS",
                GetEtabsModelPath,
                GetEtabsModelCurrentUnit,
                csiObject => csiObject.ApplicationExit(false));

            return adapter;
        }

        internal static ICSISapModelConnectionAdapter<SAP2000v1.cSapModel> CreateSap2000()
        {
            return CreateSap2000(new Sap2000ModelAdapter());
        }

        internal static ICSISapModelConnectionAdapter<SAP2000v1.cSapModel> CreateSap2000(ICsiModelAdapter modelAdapter)
        {
            var adapter = new CSISapModelConnectionAdapter<SAP2000v1.cSapModel, SAP2000v1.cOAPI>(
                modelAdapter,
                "SAP2000",
                GetSap2000ModelPath,
                GetSap2000ModelCurrentUnit,
                csiObject => csiObject.ApplicationExit(false));

            return adapter;
        }

        private static string GetEtabsModelPath(ETABSv1.cSapModel sapModel)
        {
            try
            {
                return sapModel.GetModelFilename(true);
            }
            catch
            {
                return string.Empty;
            }
        }

        private static string GetSap2000ModelPath(SAP2000v1.cSapModel sapModel)
        {
            try
            {
                return sapModel.GetModelFilename(true);
            }
            catch
            {
                return string.Empty;
            }
        }

        private static string GetEtabsModelCurrentUnit(ETABSv1.cSapModel sapModel)
        {
            try
            {
                ETABSv1.eForce forceUnits = ETABSv1.eForce.N;
                ETABSv1.eLength lengthUnits = ETABSv1.eLength.m;
                ETABSv1.eTemperature temperatureUnits = ETABSv1.eTemperature.C;

                int getUnitsResult = sapModel.GetDatabaseUnits_2(ref forceUnits, ref lengthUnits, ref temperatureUnits);
                return getUnitsResult == 0
                    ? EtabsUnitFormatter.FormatDatabaseUnits(forceUnits, lengthUnits, temperatureUnits)
                    : "Units unavailable";
            }
            catch
            {
                return "Units unavailable";
            }
        }

        private static string GetSap2000ModelCurrentUnit(SAP2000v1.cSapModel sapModel)
        {
            try
            {
                return Sap2000UnitFormatter.FormatPresentUnits(sapModel.GetPresentUnits());
            }
            catch
            {
                return "Units unavailable";
            }
        }
    }
}
