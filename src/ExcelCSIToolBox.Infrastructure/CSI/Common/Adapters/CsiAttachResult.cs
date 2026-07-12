namespace ExcelCSIToolBox.Infrastructure.CSI.Common.Adapters
{
    public class CsiAttachResult
    {
        public bool IsSuccess { get; set; }

        public object ApplicationObject { get; set; }

        public object SapModel { get; set; }

        public string InstanceId { get; set; }

        public int? ProcessId { get; set; }

        public string Message { get; set; }

        public static CsiAttachResult Success(object applicationObject, object sapModel, string message = null, string instanceId = null, int? processId = null)
        {
            var result = new CsiAttachResult
            {
                IsSuccess = true,
                ApplicationObject = applicationObject,
                SapModel = sapModel,
                InstanceId = instanceId,
                ProcessId = processId,
                Message = message
            };
            return result;
        }

        public static CsiAttachResult Failure(string message)
        {
            var result = new CsiAttachResult
            {
                IsSuccess = false,
                Message = message
            };
            return result;
        }
    }
}

