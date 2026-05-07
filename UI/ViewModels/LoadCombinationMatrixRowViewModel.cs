using System.Collections.Generic;
using ExcelCSIToolBox.Data.DTOs.CSI;

namespace ExcelCSIToolBoxAddIn.UI.ViewModels
{
    public class LoadCombinationMatrixRowViewModel : ViewModelBase
    {
        private string _loadCombinationName;
        private int _combinationType;
        private readonly Dictionary<string, string> _factorTexts = new Dictionary<string, string>(System.StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, int> _factorCaseTypes = new Dictionary<string, int>(System.StringComparer.OrdinalIgnoreCase);

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

        public int GetFactorCaseType(string loadPatternName)
        {
            if (loadPatternName == null)
            {
                return 0;
            }

            return _factorCaseTypes.TryGetValue(loadPatternName, out int caseType) ? caseType : 0;
        }

        public void SetFactorCaseType(string loadPatternName, int caseType)
        {
            if (loadPatternName == null)
            {
                return;
            }

            _factorCaseTypes[loadPatternName] = caseType;
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
                if (dto.FactorCaseTypes != null && dto.FactorCaseTypes.TryGetValue(patternName, out int caseType))
                {
                    row.SetFactorCaseType(patternName, caseType);
                }

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
