using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace ExcelCSIToolBoxAddIn.UI.ViewModels
{
    public abstract class ViewModelBase : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        public event System.EventHandler RequestHide;
        public event System.EventHandler RequestShow;

        protected void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        public void RaiseRequestHide()
        {
            RequestHide?.Invoke(this, System.EventArgs.Empty);
        }

        public void RaiseRequestShow()
        {
            RequestShow?.Invoke(this, System.EventArgs.Empty);
        }
    }
}

