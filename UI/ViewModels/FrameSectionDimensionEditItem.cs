namespace ExcelCSIToolBoxAddIn.UI.ViewModels
{
    public class FrameSectionDimensionEditItem : ViewModelBase
    {
        private string _valueText;

        public string Key { get; set; }

        public string ValueText
        {
            get { return _valueText; }
            set
            {
                _valueText = value;
                OnPropertyChanged();
            }
        }
    }
}
