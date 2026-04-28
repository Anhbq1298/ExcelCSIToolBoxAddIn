namespace ExcelCSIToolBoxAddIn.Adapters
{
    public class EtabsModelAdapter : ICsiModelAdapter
    {
        private const string EtabsComProgId = "CSI.ETABS.API.ETABSObject";

        public string ApplicationName => "ETABS";

        public CsiAttachResult AttachToRunningInstance()
        {
            ETABSv1.cHelper helper = new ETABSv1.Helper();

            try
            {
                ETABSv1.cOAPI etabsObject = helper.GetObject(EtabsComProgId);
                if (etabsObject == null)
                {
                    return CsiAttachResult.Failure("ETABS is not running.");
                }

                ETABSv1.cSapModel sapModel = etabsObject.SapModel;
                if (sapModel == null)
                {
                    return CsiAttachResult.Failure("ETABS is running, but no active SapModel could be retrieved.");
                }

                return CsiAttachResult.Success(etabsObject, sapModel, "Successfully attached to ETABS.");
            }
            catch
            {
                return CsiAttachResult.Failure("ETABS is not running.");
            }
        }
    }
}
