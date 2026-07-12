using System;
using ExcelCSIToolBox.Core.Abstractions.CSI;

namespace ExcelCSIToolBox.Infrastructure.CSI.Common
{
    public sealed class CurrentThreadCsiApiDispatcher : ICsiApiDispatcher
    {
        public void Invoke(Action operation)
        {
            if (operation == null)
            {
                throw new ArgumentNullException(nameof(operation));
            }

            operation();
        }

        public T Invoke<T>(Func<T> operation)
        {
            if (operation == null)
            {
                throw new ArgumentNullException(nameof(operation));
            }

            return operation();
        }
    }
}
