using System;

namespace ExcelCSIToolBox.Core.Abstractions.CSI
{
    public interface ICsiApiDispatcher
    {
        void Invoke(Action operation);

        T Invoke<T>(Func<T> operation);
    }
}
