using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace CsiExcelAddin.ViewModels
{
    /// <summary>
    /// Base class for all ViewModels in this add-in.
    /// Implements INotifyPropertyChanged so WPF bindings update automatically.
    /// All ViewModels must inherit from this â€” do not reimplement INPC elsewhere.
    /// </summary>
    public abstract class BaseViewModel : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        /// <summary>
        /// Sets the backing field and raises PropertyChanged only when the value
        /// has actually changed. This prevents unnecessary UI redraws.
        /// CallerMemberName removes the need to pass the property name manually.
        /// </summary>
        protected bool SetProperty<T>(ref T field, T value, [CallerMemberName] string propertyName = null)
        {
            if (Equals(field, value)) return false;
            field = value;
            OnPropertyChanged(propertyName);
            return true;
        }

        /// <summary>
        /// Raises PropertyChanged for the given property name.
        /// Call with nameof(PropertyName) when manual notification is needed.
        /// </summary>
        protected void OnPropertyChanged([CallerMemberName] string propertyName = null)
            => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }
}

