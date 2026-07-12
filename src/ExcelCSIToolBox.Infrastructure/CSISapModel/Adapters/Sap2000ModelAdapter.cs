namespace ExcelCSIToolBox.Infrastructure.CSISapModel.Adapters
{
    using System.Collections.Generic;

    public class Sap2000ModelAdapter : ICsiModelAdapter
    {
        private const string Sap2000ComProgId = "CSI.SAP2000.API.SapObject";

        public string ApplicationName => "SAP2000";

        public IReadOnlyList<CsiAttachResult> GetRunningInstances()
        {
            var instances = new List<CsiAttachResult>(
                CsiRunningObjectTable.GetRunningInstances<SAP2000v1.cOAPI, SAP2000v1.cSapModel>(
                    ApplicationName,
                    Sap2000ComProgId,
                    sapObject => sapObject.SapModel));

            if (instances.Count == 0)
            {
                CsiAttachResult attachResult = AttachToRunningInstance();
                if (attachResult.IsSuccess)
                {
                    attachResult.InstanceId = string.IsNullOrWhiteSpace(attachResult.InstanceId)
                        ? Sap2000ComProgId
                        : attachResult.InstanceId;
                    instances.Add(attachResult);
                }
            }

            return instances;
        }

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

                return CsiAttachResult.Success(sapObject, sapModel, "Successfully attached to SAP2000.", Sap2000ComProgId);
            }
            catch
            {
                return CsiAttachResult.Failure("SAP2000 is not running.");
            }
        }

        public CsiAttachResult AttachToRunningInstance(string instanceId)
        {
            if (string.IsNullOrWhiteSpace(instanceId) || string.Equals(instanceId, Sap2000ComProgId, System.StringComparison.OrdinalIgnoreCase))
            {
                return AttachToRunningInstance();
            }

            foreach (CsiAttachResult instance in GetRunningInstances())
            {
                if (string.Equals(instance.InstanceId, instanceId, System.StringComparison.OrdinalIgnoreCase))
                {
                    instance.Message = "Successfully attached to SAP2000.";
                    return instance;
                }
            }

            return CsiAttachResult.Failure("The selected SAP2000 instance is no longer running.");
        }
    }
}

