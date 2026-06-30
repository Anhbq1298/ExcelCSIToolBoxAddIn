namespace ExcelCSIToolBox.Infrastructure.CSISapModel.Adapters
{
    using System.Collections.Generic;

    public class EtabsModelAdapter : ICsiModelAdapter
    {
        private const string EtabsComProgId = "CSI.ETABS.API.ETABSObject";

        public string ApplicationName => "ETABS";

        public IReadOnlyList<CsiAttachResult> GetRunningInstances()
        {
            var instances = new List<CsiAttachResult>(
                CsiRunningObjectTable.GetRunningInstances<ETABSv1.cOAPI, ETABSv1.cSapModel>(
                    ApplicationName,
                    EtabsComProgId,
                    etabsObject => etabsObject.SapModel));

            if (instances.Count == 0)
            {
                CsiAttachResult attachResult = AttachToRunningInstance();
                if (attachResult.IsSuccess)
                {
                    attachResult.InstanceId = string.IsNullOrWhiteSpace(attachResult.InstanceId)
                        ? EtabsComProgId
                        : attachResult.InstanceId;
                    instances.Add(attachResult);
                }
            }

            return instances;
        }

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

                return CsiAttachResult.Success(etabsObject, sapModel, "Successfully attached to ETABS.", EtabsComProgId);
            }
            catch
            {
                return CsiAttachResult.Failure("ETABS is not running.");
            }
        }

        public CsiAttachResult AttachToRunningInstance(string instanceId)
        {
            if (string.IsNullOrWhiteSpace(instanceId) || string.Equals(instanceId, EtabsComProgId, System.StringComparison.OrdinalIgnoreCase))
            {
                return AttachToRunningInstance();
            }

            foreach (CsiAttachResult instance in GetRunningInstances())
            {
                if (string.Equals(instance.InstanceId, instanceId, System.StringComparison.OrdinalIgnoreCase))
                {
                    instance.Message = "Successfully attached to ETABS.";
                    return instance;
                }
            }

            return CsiAttachResult.Failure("The selected ETABS instance is no longer running.");
        }
    }
}

