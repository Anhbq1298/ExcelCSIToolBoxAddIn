namespace ExcelCSIToolBox.Core.Common.Results
{
    /// <summary>
    /// Simple operation result wrapper for safe error handling without exceptions in UI flow.
    /// </summary>
    public class OperationResult
    {
        public bool IsSuccess { get; protected set; }

        public string Message { get; protected set; }

        public static OperationResult Success(string message = null)
        {
            var result = new OperationResult { IsSuccess = true, Message = message };
            return result;
        }

        public static OperationResult Failure(string message)
        {
            var result = new OperationResult { IsSuccess = false, Message = message };
            return result;
        }
    }

    public class OperationResult<T> : OperationResult
    {
        public T Data { get; private set; }

        public static OperationResult<T> Success(T data, string message = null)
        {
            var result = new OperationResult<T> { IsSuccess = true, Data = data, Message = message };
            return result;
        }

        public static new OperationResult<T> Failure(string message)
        {
            var result = new OperationResult<T> { IsSuccess = false, Message = message };
            return result;
        }
    }
}

