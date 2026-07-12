using ExcelCSIToolBox.Data.DTOs.CSI;

namespace ExcelCSIToolBoxAddIn.UI.ViewModels
{
    public class BaseReactionOutputCaseViewModel : ViewModelBase
    {
        private bool _isSelected;

        public BaseReactionOutputCaseViewModel(CSISapModelOutputCaseDTO dto)
        {
            OutputCase = dto;
        }

        public CSISapModelOutputCaseDTO OutputCase { get; private set; }

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

        public string Name
        {
            get { return OutputCase == null ? string.Empty : OutputCase.Name; }
        }

        public string Kind
        {
            get { return OutputCase == null ? string.Empty : OutputCase.Kind; }
        }

        public string Type
        {
            get { return OutputCase == null ? string.Empty : OutputCase.Type; }
        }
    }
}
