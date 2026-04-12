using System;
using System.IO;
using ExcelCSIToolBoxAddIn.Common.Results;

namespace ExcelCSIToolBoxAddIn.Infrastructure.Etabs
{
    /// <summary>
    /// Infrastructure service that safely attaches to a running ETABS instance.
    /// </summary>
    public class EtabsConnectionService : IEtabsConnectionService
    {
        private const string EtabsComProgId = "CSI.ETABS.API.ETABSObject";

        public OperationResult<EtabsConnectionInfo> TryAttachToRunningInstance()
        {
            // create API helper object
            ETABSv1.cHelper myHelper = new ETABSv1.Helper();
            ETABSv1.cOAPI myETABSObject = null;

            try
            {
                // attach to a running instance of ETABS
                myETABSObject = myHelper.GetObject(EtabsComProgId);

                if (myETABSObject == null)
                {
                    return OperationResult<EtabsConnectionInfo>.Failure("ETABS is not running.");
                }

                ETABSv1.cSapModel sapModel = myETABSObject.SapModel;
                string modelPath = GetModelFileNameSafely(sapModel);
                string modelName = string.IsNullOrWhiteSpace(modelPath)
                    ? "Unknown model"
                    : Path.GetFileName(modelPath);

                return OperationResult<EtabsConnectionInfo>.Success(new EtabsConnectionInfo
                {
                    IsConnected = true,
                    ModelFileName = modelName,
                    EtabsObject = myETABSObject,
                    SapModel = sapModel
                });
            }
            catch (Exception)
            {
                return OperationResult<EtabsConnectionInfo>.Failure("ETABS is not running.");
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
