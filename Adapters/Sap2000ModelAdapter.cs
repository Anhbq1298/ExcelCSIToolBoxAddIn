namespace ExcelCSIToolBoxAddIn.Adapters
{
    public class Sap2000ModelAdapter : ICsiModelAdapter
    {
        private const string Sap2000ComProgId = "CSI.SAP2000.API.SapObject";

        public string ApplicationName => "SAP2000";

        public CsiAttachResult AttachToRunningInstance()
        {
            SAP2000v1.cHelper helper = new SAP2000v1.Helper();

            try
            {
                SAP2000v1.cOAPI sapObject = helper.GetObject(Sap2000ComProgId);
                if (sapObject == null)
                {
                    return CsiAttachResult.Failure("SAP2000 is not running.");
                }

                SAP2000v1.cSapModel sapModel = sapObject.SapModel;
                if (sapModel == null)
                {
                    return CsiAttachResult.Failure("SAP2000 is running, but no active SapModel could be retrieved.");
                }

                return CsiAttachResult.Success(sapObject, sapModel, "Successfully attached to SAP2000.");
            }
            catch
            {
                return CsiAttachResult.Failure("SAP2000 is not running.");
            }
        }
    }
}
