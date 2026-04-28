using System;
using System.Windows.Input;

namespace CsiExcelAddin.Commands
{
    /// <summary>
    /// Standard relay command that delegates execution and can-execute logic
    /// to caller-supplied delegates. Use this for synchronous ViewModel actions.
    /// </summary>
    public class RelayCommand : ICommand
    {
        private readonly Action<object> _execute;
        private readonly Func<object, bool> _canExecute;

        /// <summary>
        /// Raised by WPF command infrastructure to re-evaluate CanExecute.
        /// Hooked to CommandManager so UI refreshes automatically.
        /// </summary>
        public event EventHandler CanExecuteChanged
        {
            add => CommandManager.RequerySuggested += value;
            remove => CommandManager.RequerySuggested -= value;
        }

        /// <param name="execute">Action to run when command is invoked.</param>
        /// <param name="canExecute">Optional guard â€” returns true when command is allowed.</param>
        public RelayCommand(Action<object> execute, Func<object, bool> canExecute = null)
        {
            _execute = execute ?? throw new ArgumentNullException(nameof(execute));
            _canExecute = canExecute;
        }

        /// <summary>Convenience constructor for commands with no parameter.</summary>
        public RelayCommand(Action execute, Func<bool> canExecute = null)
            : this(_ => execute(), canExecute is null ? null : _ => canExecute())
        {
        }

        public bool CanExecute(object parameter) => _canExecute?.Invoke(parameter) ?? true;

        public void Execute(object parameter) => _execute(parameter);

        /// <summary>
        /// Forces WPF to re-query CanExecute. Call this when the ViewModel state
        /// changes and the button enabled state must update immediately.
        /// </summary>
        public void RaiseCanExecuteChanged()
            => CommandManager.InvalidateRequerySuggested();
    }
}

