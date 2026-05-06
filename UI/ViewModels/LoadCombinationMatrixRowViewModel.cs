using System.Collections.Generic;
using ExcelCSIToolBox.Data.DTOs.CSI;

namespace ExcelCSIToolBoxAddIn.UI.ViewModels
{
    public class LoadCombinationMatrixRowViewModel : ViewModelBase
    {
        private string _loadCombinationName;
        private int _combinationType;
        private readonly Dictionary<string, string> _factorTexts = new Dictionary<string, string>(System.StringComparer.OrdinalIgnoreCase);

        public string LoadCombinationName
        {
            get { return _loadCombinationName; }
            set
            {
                _loadCombinationName = value;
                OnPropertyChanged();
            }
        }

        public int CombinationType
        {
            get { return _combinationType; }
            set
            {
                _combinationType = value;
                OnPropertyChanged();
            }
        }

        public string this[string loadPatternName]
        {
            get
            {
                if (loadPatternName == null)
                {
                    return null;
                }

                return _factorTexts.TryGetValue(loadPatternName, out string value) ? value : null;
            }
            set
            {
                if (loadPatternName == null)
                {
                    return;
                }

                _factorTexts[loadPatternName] = value;
                OnPropertyChanged("Item[]");
            }
        }

        public static LoadCombinationMatrixRowViewModel FromDto(LoadCombinationMatrixRowDto dto, System.Collections.Generic.IEnumerable<string> loadPatternNames)
        {
            var row = new LoadCombinationMatrixRowViewModel
            {
                LoadCombinationName = dto.LoadCombinationName,
                CombinationType = dto.CombinationType
            };

            foreach (string patternName in loadPatternNames)
            {
                if (dto.Factors != null && dto.Factors.TryGetValue(patternName, out double? factor) && factor.HasValue)
                {
                    row[patternName] = factor.Value.ToString("G15", System.Globalization.CultureInfo.InvariantCulture);
                }
                else
                {
                    row[patternName] = null;
                }
            }

            return row;
        }
    }
}
