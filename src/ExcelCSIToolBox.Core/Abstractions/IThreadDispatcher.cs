using System;
using System.Threading.Tasks;

namespace ExcelCSIToolBox.Core.Abstractions
{
    /// <summary>
    /// Marshals work back to the UI thread that owns Excel COM and WPF objects.
    /// </summary>
    public interface IThreadDispatcher
    {
        void InvokeOnUiThread(Action action);

        Task InvokeOnUiThreadAsync(Func<Task> action);
    }
}
