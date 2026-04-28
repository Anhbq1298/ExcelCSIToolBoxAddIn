using System.Collections.Generic;

namespace ExcelCSIToolBox.Core.Common.Results
{
    public class CsiOperationResult<T>
    {
        public bool Success { get; set; }
        public string Message { get; set; }
        public T Data { get; set; }
        public List<string> Warnings { get; set; }
        public List<string> Errors { get; set; }

        public CsiOperationResult()
        {
            Warnings = new List<string>();
            Errors = new List<string>();
        }

        public static CsiOperationResult<T> Ok(T data, string message)
        {
            return new CsiOperationResult<T>
            {
                Success = true,
                Message = message,
                Data = data
            };
        }

        public static CsiOperationResult<T> Fail(string message)
        {
            return Fail(message, new List<string>());
        }

        public static CsiOperationResult<T> Fail(string message, List<string> errors)
        {
            return new CsiOperationResult<T>
            {
                Success = false,
                Message = message,
                Errors = errors ?? new List<string>()
            };
        }

        public static CsiOperationResult<T> OkWithWarnings(T data, string message, List<string> warnings)
        {
            return new CsiOperationResult<T>
            {
                Success = true,
                Message = message,
                Data = data,
                Warnings = warnings ?? new List<string>()
            };
        }
    }
}
