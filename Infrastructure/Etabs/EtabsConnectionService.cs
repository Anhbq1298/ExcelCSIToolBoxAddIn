using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using ExcelCSIToolBoxAddIn.Common.Results;

namespace ExcelCSIToolBoxAddIn.Infrastructure.Etabs
{
    /// <summary>
    /// Infrastructure service that safely attaches to a running ETABS instance.
    /// Uses late binding so phase 1 can run without adding ETABS interop assemblies.
    /// </summary>
    public class EtabsConnectionService : IEtabsConnectionService
    {
        private const string EtabsComProgId = "CSI.ETABS.API.ETABSObject";

        public OperationResult<EtabsConnectionInfo> TryAttachToRunningInstance()
        {
            try
            {
                var etabsObject = Marshal.GetActiveObject(EtabsComProgId);
                if (etabsObject == null)
                {
                    return OperationResult<EtabsConnectionInfo>.Failure("Could not find a running ETABS instance.");
                }

                var sapModel = GetPropertyValue(etabsObject, "SapModel");
                var modelPath = GetModelFileNameSafely(sapModel);
                var fileName = string.IsNullOrWhiteSpace(modelPath) ? "Unknown model" : Path.GetFileName(modelPath);

                return OperationResult<EtabsConnectionInfo>.Success(new EtabsConnectionInfo
                {
                    IsConnected = true,
                    ModelFileName = fileName,
                    EtabsObject = etabsObject,
                    SapModel = sapModel
                });
            }
            catch (COMException)
            {
                return OperationResult<EtabsConnectionInfo>.Failure("ETABS is not running.");
            }
            catch (Exception ex)
            {
                return OperationResult<EtabsConnectionInfo>.Failure($"Unable to connect to ETABS: {ex.Message}");
            }
        }

        private static object GetPropertyValue(object target, string propertyName)
        {
            if (target == null)
            {
                return null;
            }

            return target.GetType().InvokeMember(
                propertyName,
                BindingFlags.GetProperty,
                null,
                target,
                null);
        }

        private static string GetModelFileNameSafely(object sapModel)
        {
            if (sapModel == null)
            {
                return null;
            }

            try
            {
                // ETABS API signatures vary by version. Try common overloads.
                var modelPath = sapModel.GetType().InvokeMember(
                    "GetModelFilename",
                    BindingFlags.InvokeMethod,
                    null,
                    sapModel,
                    new object[] { true });

                return modelPath as string;
            }
            catch
            {
                try
                {
                    var modelPath = sapModel.GetType().InvokeMember(
                        "GetModelFilename",
                        BindingFlags.InvokeMethod,
                        null,
                        sapModel,
                        null);

                    return modelPath as string;
                }
                catch
                {
                    return null;
                }
            }
        }
    }
}
