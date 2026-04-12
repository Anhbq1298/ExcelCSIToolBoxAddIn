namespace ExcelCSIToolBoxAddIn.Common.Results
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
            return new OperationResult { IsSuccess = true, Message = message };
        }

        public static OperationResult Failure(string message)
        {
            return new OperationResult { IsSuccess = false, Message = message };
        }
    }

    public class OperationResult<T> : OperationResult
    {
        public T Data { get; private set; }

        public static OperationResult<T> Success(T data, string message = null)
        {
            return new OperationResult<T> { IsSuccess = true, Data = data, Message = message };
        }

        public static new OperationResult<T> Failure(string message)
        {
            return new OperationResult<T> { IsSuccess = false, Message = message };
        }
    }
}
