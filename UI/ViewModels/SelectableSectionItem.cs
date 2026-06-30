namespace ExcelCSIToolBoxAddIn.UI.ViewModels
{
    public class SelectableSectionItem : ViewModelBase
    {
        private bool _isSelected;

        public SelectableSectionItem(string name)
        {
            Name = name;
        }

        public string Name { get; private set; }

        public bool IsSelected
        {
            get { return _isSelected; }
            set
            {
                if (_isSelected == value)
                {
                    return;
                }

                _isSelected = value;
                OnPropertyChanged();
            }
        }
    }
}
