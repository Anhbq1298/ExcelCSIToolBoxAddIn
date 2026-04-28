namespace ExcelCSIToolBox.Infrastructure.CSISapModel.Adapters
{
    public class CsiAttachResult
    {
        public bool IsSuccess { get; set; }

        public object ApplicationObject { get; set; }

        public object SapModel { get; set; }

        public string Message { get; set; }

        public static CsiAttachResult Success(object applicationObject, object sapModel, string message = null)
        {
            var result = new CsiAttachResult
            {
                IsSuccess = true,
                ApplicationObject = applicationObject,
                SapModel = sapModel,
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

