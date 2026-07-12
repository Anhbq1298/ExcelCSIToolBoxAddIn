using System;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Threading;
using ExcelCSIToolBox.Core.Abstractions;

namespace ExcelCSIToolBoxAddIn.AddIn
{
    /// <summary>
    /// WPF dispatcher implementation for marshaling Excel/WPF work to the STA UI thread.
    /// </summary>
    internal sealed class WpfThreadDispatcher : IThreadDispatcher
    {
        public void InvokeOnUiThread(Action action)
        {
            if (action == null)
            {
                return;
            }

            Dispatcher dispatcher = GetDispatcher();
            if (dispatcher == null || dispatcher.CheckAccess())
            {
                action();
                return;
            }

            dispatcher.Invoke(action);
        }

        public Task InvokeOnUiThreadAsync(Func<Task> action)
        {
            if (action == null)
            {
                return Task.FromResult<object>(null);
            }

            Dispatcher dispatcher = GetDispatcher();
            if (dispatcher == null || dispatcher.CheckAccess())
            {
                return action();
            }

            return dispatcher.Invoke(action);
        }

        private static Dispatcher GetDispatcher()
        {
            return Application.Current == null ? null : Application.Current.Dispatcher;
        }
    }
}
