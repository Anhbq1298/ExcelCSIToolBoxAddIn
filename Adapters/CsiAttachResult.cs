namespace ExcelCSIToolBoxAddIn.Adapters
{
    public class CsiAttachResult
    {
        public bool IsSuccess { get; set; }

        public object ApplicationObject { get; set; }

        public object SapModel { get; set; }

        public string Message { get; set; }

        public static CsiAttachResult Success(object applicationObject, object sapModel, string message = null)
        {
            return new CsiAttachResult
            {
                IsSuccess = true,
                ApplicationObject = applicationObject,
                SapModel = sapModel,
                Message = message
            };
        }

        public static CsiAttachResult Failure(string message)
        {
            return new CsiAttachResult
            {
                IsSuccess = false,
                Message = message
            };
        }
    }
}
