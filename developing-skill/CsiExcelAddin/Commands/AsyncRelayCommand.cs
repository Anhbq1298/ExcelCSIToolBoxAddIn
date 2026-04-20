using System;
using System.Threading.Tasks;
using System.Windows.Input;

namespace CsiExcelAddin.Commands
{
    /// <summary>
    /// Async relay command for long-running operations such as CSI API calls
    /// or Excel data reads. Prevents re-entry while execution is in progress.
    /// </summary>
    public class AsyncRelayCommand : ICommand
    {
        private readonly Func<object, Task> _execute;
        private readonly Func<object, bool> _canExecute;
        private bool _isExecuting;

        public event EventHandler CanExecuteChanged
        {
            add => CommandManager.RequerySuggested += value;
            remove => CommandManager.RequerySuggested -= value;
        }

        /// <param name="execute">Async action to run when command is invoked.</param>
        /// <param name="canExecute">Optional guard — also blocks when already executing.</param>
        public AsyncRelayCommand(Func<object, Task> execute, Func<object, bool> canExecute = null)
        {
            _execute = execute ?? throw new ArgumentNullException(nameof(execute));
            _canExecute = canExecute;
        }

        /// <summary>Convenience constructor for async commands with no parameter.</summary>
        public AsyncRelayCommand(Func<Task> execute, Func<bool> canExecute = null)
            : this(_ => execute(), canExecute is null ? null : _ => canExecute())
        {
        }

        /// <summary>
        /// Returns false while an execution is in progress to prevent re-entry.
        /// </summary>
        public bool CanExecute(object parameter)
            => !_isExecuting && (_canExecute?.Invoke(parameter) ?? true);

        public async void Execute(object parameter)
        {
            if (!CanExecute(parameter)) return;

            _isExecuting = true;
            CommandManager.InvalidateRequerySuggested();

            try
            {
                await _execute(parameter);
            }
            finally
            {
                // Always restore state so the button re-enables even on failure
                _isExecuting = false;
                CommandManager.InvalidateRequerySuggested();
            }
        }
    }
}
