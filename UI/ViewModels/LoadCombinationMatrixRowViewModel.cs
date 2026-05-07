using System.Collections.Generic;
using ExcelCSIToolBox.Data.DTOs.CSI;

namespace ExcelCSIToolBoxAddIn.UI.ViewModels
{
    public class LoadCombinationMatrixRowViewModel : ViewModelBase
    {
        private string _loadCombinationName;
        private int _combinationType;
        private readonly Dictionary<string, string> _loadCaseFactorTexts = new Dictionary<string, string>(System.StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, string> _loadCombinationFactorTexts = new Dictionary<string, string>(System.StringComparer.OrdinalIgnoreCase);
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

        public Dictionary<string, string> LoadCaseFactors
        {
            get { return _loadCaseFactorTexts; }
        }

        public Dictionary<string, string> LoadCombinationFactors
        {
            get { return _loadCombinationFactorTexts; }
        }

        public string this[string loadPatternName]
        {
            get
            {
                if (loadPatternName == null)
                {
                    return null;
                }

                return GetLoadCaseFactor(loadPatternName);
            }
            set
            {
                SetLoadCaseFactor(loadPatternName, value);
            }
        }

        public string GetLoadCaseFactor(string loadPatternName)
        {
            if (loadPatternName == null)
            {
                return null;
            }

            return _loadCaseFactorTexts.TryGetValue(loadPatternName, out string value) ? value : null;
        }

        public void SetLoadCaseFactor(string loadPatternName, string value)
        {
            if (loadPatternName == null)
            {
                return;
            }

            _loadCaseFactorTexts[loadPatternName] = value;
            OnPropertyChanged("Item[]");
            OnPropertyChanged(nameof(LoadCaseFactors));
        }

        public string GetLoadCombinationFactor(string loadCombinationName)
        {
            if (loadCombinationName == null)
            {
                return null;
            }

            return _loadCombinationFactorTexts.TryGetValue(loadCombinationName, out string value) ? value : null;
        }

        public void SetLoadCombinationFactor(string loadCombinationName, string value)
        {
            if (loadCombinationName == null)
            {
                return;
            }

            _loadCombinationFactorTexts[loadCombinationName] = value;
            OnPropertyChanged(nameof(LoadCombinationFactors));
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

        public static LoadCombinationMatrixRowViewModel FromDto(
            LoadCombinationMatrixRowDto dto,
            System.Collections.Generic.IEnumerable<string> loadPatternNames,
            System.Collections.Generic.IEnumerable<string> loadCombinationReferenceNames)
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

                double? factor = null;
                if (dto.LoadCaseFactors != null && dto.LoadCaseFactors.TryGetValue(patternName, out double? loadCaseFactor))
                {
                    factor = loadCaseFactor;
                }
                else if (dto.Factors != null && dto.Factors.TryGetValue(patternName, out double? fallbackFactor))
                {
                    factor = fallbackFactor;
                }

                if (factor.HasValue)
                {
                    row.SetLoadCaseFactor(patternName, factor.Value.ToString("G15", System.Globalization.CultureInfo.InvariantCulture));
                }
                else
                {
                    row.SetLoadCaseFactor(patternName, null);
                }
            }

            foreach (string comboName in loadCombinationReferenceNames)
            {
                if (dto.LoadCombinationFactors != null && dto.LoadCombinationFactors.TryGetValue(comboName, out double? factor) && factor.HasValue)
                {
                    row.SetLoadCombinationFactor(
                        comboName,
                        factor.Value.ToString("G15", System.Globalization.CultureInfo.InvariantCulture));
                }
                else
                {
                    row.SetLoadCombinationFactor(comboName, null);
                }
            }

            return row;
        }
    }
}
