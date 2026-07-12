namespace ExcelCSIToolBoxAddIn.UI.ViewModels
{
    public class ModifierFieldViewModel : ViewModelBase
    {
        private string _valueText;

        public ModifierFieldViewModel(string label)
        {
            Label = label;
            ValueText = "1";
        }

        public string Label { get; private set; }

        public string ValueText
        {
            get { return _valueText; }
            set
            {
                if (_valueText == value)
                {
                    return;
                }

                _valueText = value;
                OnPropertyChanged();
            }
        }
    }
}
