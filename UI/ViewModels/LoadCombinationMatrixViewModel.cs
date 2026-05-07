using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;
using System.Windows;
using System.Windows.Input;
using ExcelCSIToolBox.Core.Common.Commands;
using ExcelCSIToolBox.Core.Common.Results;
using ExcelCSIToolBox.Core.Abstractions.Excel;
using ExcelCSIToolBox.Data.DTOs.CSI;

namespace ExcelCSIToolBoxAddIn.UI.ViewModels
{
    public class LoadCombinationMatrixViewModel : ViewModelBase
    {
        private readonly string _productTitle;
        private readonly IExcelOutputService _excelOutputService;
        private readonly HashSet<string> _originalLoadCombinationNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private IReadOnlyList<LoadCombinationMatrixRowDto> _savedRows = new List<LoadCombinationMatrixRowDto>();
        private IReadOnlyList<string> _savedDeletedLoadCombinationNames = new List<string>();

        public LoadCombinationMatrixViewModel(
            LoadCombinationMatrixDto initialMatrix,
            string productTitle,
            IExcelOutputService excelOutputService = null)
        {
            _productTitle = string.IsNullOrWhiteSpace(productTitle) ? "CSI Toolbox" : productTitle;
            _excelOutputService = excelOutputService;

            Rows = new ObservableCollection<LoadCombinationMatrixRowViewModel>();
            LoadPatternNames = new ObservableCollection<string>();
            LoadCombinationReferenceNames = new ObservableCollection<string>();
            CombinationTypeOptions = CreateCombinationTypeOptions();

            AddRowCommand = new RelayCommand(AddRow);
            DeleteSelectedRowsCommand = new RelayCommand<IList>(DeleteSelectedRows);
            ExportToExcelRangeCommand = new RelayCommand(ExportToExcelRange, () => _excelOutputService != null);
            SaveCommand = new RelayCommand(Save);
            CancelCommand = new RelayCommand(Cancel);

            LoadMatrix(initialMatrix ?? new LoadCombinationMatrixDto());
        }

        public ObservableCollection<LoadCombinationMatrixRowViewModel> Rows { get; }
        public ObservableCollection<string> LoadPatternNames { get; }
        public ObservableCollection<string> LoadCombinationReferenceNames { get; }
        public ObservableCollection<LoadCombinationTypeOption> CombinationTypeOptions { get; }

        public ICommand AddRowCommand { get; }
        public ICommand DeleteSelectedRowsCommand { get; }
        public ICommand ExportToExcelRangeCommand { get; }
        public ICommand SaveCommand { get; }
        public ICommand CancelCommand { get; }

        public bool WasSaved { get; private set; }

        public IReadOnlyList<LoadCombinationMatrixRowDto> SavedRows
        {
            get { return _savedRows; }
        }

        public IReadOnlyList<string> SavedDeletedLoadCombinationNames
        {
            get { return _savedDeletedLoadCombinationNames; }
        }

        public event EventHandler RequestClose;
        public event EventHandler RequestExcelSelectionStart;
        public event EventHandler RequestExcelSelectionEnd;

        private void LoadMatrix(LoadCombinationMatrixDto matrix)
        {
            _originalLoadCombinationNames.Clear();
            LoadPatternNames.Clear();
            LoadCombinationReferenceNames.Clear();
            Rows.Clear();

            var patternNames = matrix.LoadPatternNames ?? new List<string>();
            foreach (string patternName in patternNames)
            {
                string trimmed = string.IsNullOrWhiteSpace(patternName) ? null : patternName.Trim();
                if (!string.IsNullOrWhiteSpace(trimmed) &&
                    !LoadPatternNames.Any(x => string.Equals(x, trimmed, StringComparison.OrdinalIgnoreCase)))
                {
                    LoadPatternNames.Add(trimmed);
                }
            }

            var referenceNames = matrix.LoadCombinationReferenceNames ?? new List<string>();
            foreach (string referenceName in referenceNames)
            {
                string trimmed = NormalizeName(referenceName);
                if (!string.IsNullOrWhiteSpace(trimmed) &&
                    !LoadCombinationReferenceNames.Any(x => string.Equals(x, trimmed, StringComparison.OrdinalIgnoreCase)))
                {
                    LoadCombinationReferenceNames.Add(trimmed);
                }
            }

            if (matrix.Rows != null)
            {
                foreach (var row in matrix.Rows)
                {
                    string originalName = NormalizeName(row.LoadCombinationName);
                    if (!string.IsNullOrWhiteSpace(originalName))
                    {
                        _originalLoadCombinationNames.Add(originalName);
                    }

                    Rows.Add(LoadCombinationMatrixRowViewModel.FromDto(row, LoadPatternNames, LoadCombinationReferenceNames));
                }
            }
        }

        private void AddRow()
        {
            var row = new LoadCombinationMatrixRowViewModel
            {
                LoadCombinationName = "New Combo",
                CombinationType = (int)LoadCombinationType.LinearAdditive
            };

            foreach (string patternName in LoadPatternNames)
            {
                row[patternName] = null;
                row.SetFactorCaseType(patternName, 0);
            }

            foreach (string comboName in LoadCombinationReferenceNames)
            {
                row.SetLoadCombinationFactor(comboName, null);
            }

            Rows.Add(row);
        }

        private void DeleteSelectedRows(IList selectedItems)
        {
            if (selectedItems == null || selectedItems.Count == 0)
            {
                return;
            }

            var selectedRows = selectedItems
                .OfType<LoadCombinationMatrixRowViewModel>()
                .Distinct()
                .ToList();

            foreach (var row in selectedRows)
            {
                Rows.Remove(row);
            }
        }

        private void ExportToExcelRange()
        {
            var valuesResult = CreateExportValues();
            if (!valuesResult.IsSuccess)
            {
                MessageBox.Show(valuesResult.Message, _productTitle, MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            OperationResult exportResult;
            RequestExcelSelectionStart?.Invoke(this, EventArgs.Empty);
            try
            {
                exportResult = _excelOutputService.WriteValuesToSelectedCell(
                    valuesResult.Data,
                    "Select the top-left cell for the load combination matrix export:",
                    "Export Load Combination Matrix",
                    $"Successfully exported {Rows.Count} load combination row(s) to Excel.");
            }
            finally
            {
                RequestExcelSelectionEnd?.Invoke(this, EventArgs.Empty);
            }

            MessageBox.Show(
                exportResult.Message,
                _productTitle,
                MessageBoxButton.OK,
                exportResult.IsSuccess ? MessageBoxImage.Information : MessageBoxImage.Warning);
        }

        private void Save()
        {
            var validation = ValidateAndCreateDtos();
            if (!validation.IsSuccess)
            {
                MessageBox.Show(validation.Message, _productTitle, MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            _savedRows = validation.Data;
            _savedDeletedLoadCombinationNames = CreateDeletedLoadCombinationNames(validation.Data);
            WasSaved = true;
            RequestClose?.Invoke(this, EventArgs.Empty);
        }

        private void Cancel()
        {
            WasSaved = false;
            RequestClose?.Invoke(this, EventArgs.Empty);
        }

        private OperationResult<IReadOnlyList<LoadCombinationMatrixRowDto>> ValidateAndCreateDtos()
        {
            var errors = new List<string>();
            var dtos = new List<LoadCombinationMatrixRowDto>();
            var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (var row in Rows)
            {
                string name = NormalizeName(row.LoadCombinationName);
                if (!string.IsNullOrWhiteSpace(name))
                {
                    names.Add(name);
                }
            }

            int rowNumber = 1;
            foreach (var row in Rows)
            {
                string name = NormalizeName(row.LoadCombinationName);
                if (string.IsNullOrWhiteSpace(name))
                {
                    errors.Add($"Row {rowNumber}: LoadCombinationName is required.");
                }
                else if (Rows.Count(x => string.Equals(NormalizeName(x.LoadCombinationName), name, StringComparison.OrdinalIgnoreCase)) > 1)
                {
                    errors.Add($"Row {rowNumber}: duplicate LoadCombinationName '{name}'.");
                }

                if (!IsSupportedCombinationType(row.CombinationType))
                {
                    errors.Add($"Row {rowNumber}: Combination Type is not supported.");
                }

                var dto = new LoadCombinationMatrixRowDto
                {
                    LoadCombinationName = name,
                    CombinationType = row.CombinationType
                };

                foreach (string patternName in LoadPatternNames)
                {
                    string factorText = row.GetLoadCaseFactor(patternName);
                    if (string.IsNullOrWhiteSpace(factorText))
                    {
                        continue;
                    }

                    if (!TryParseDouble(factorText, out double factor))
                    {
                        errors.Add($"Row {rowNumber}, column {patternName}: value '{factorText}' is not numeric.");
                        continue;
                    }

                    dto.LoadCaseFactors[patternName] = factor;
                    dto.Factors[patternName] = factor;
                    dto.FactorCaseTypes[patternName] = row.GetFactorCaseType(patternName);
                }

                foreach (string comboName in LoadCombinationReferenceNames)
                {
                    string factorText = row.GetLoadCombinationFactor(comboName);
                    if (string.IsNullOrWhiteSpace(factorText))
                    {
                        continue;
                    }

                    if (string.Equals(name, comboName, StringComparison.OrdinalIgnoreCase))
                    {
                        errors.Add($"Row {rowNumber}, column COMBO | {comboName}: a load combination cannot reference itself.");
                        continue;
                    }

                    if (!names.Contains(comboName))
                    {
                        errors.Add($"Row {rowNumber}, column COMBO | {comboName}: referenced load combination does not exist in the matrix.");
                        continue;
                    }

                    if (!TryParseDouble(factorText, out double factor))
                    {
                        errors.Add($"Row {rowNumber}, column COMBO | {comboName}: value '{factorText}' is not numeric.");
                        continue;
                    }

                    dto.LoadCombinationFactors[comboName] = factor;
                    dto.Factors[comboName] = factor;
                    dto.FactorCaseTypes[comboName] = 1;
                }

                dtos.Add(dto);
                rowNumber++;
            }

            if (dtos.Count == 0 && _originalLoadCombinationNames.Count == 0)
            {
                errors.Add("No load combination rows are available to save.");
            }

            if (errors.Count > 0)
            {
                return OperationResult<IReadOnlyList<LoadCombinationMatrixRowDto>>.Failure("Matrix validation failed: " + string.Join(" ", errors));
            }

            return OperationResult<IReadOnlyList<LoadCombinationMatrixRowDto>>.Success(dtos);
        }

        private OperationResult<object[,]> CreateExportValues()
        {
            int columnCount = 2 + LoadPatternNames.Count + LoadCombinationReferenceNames.Count;
            var values = new object[Rows.Count + 1, columnCount];

            values[0, 0] = "Load Combination Name";
            values[0, 1] = "Combination Type";

            int columnIndex = 2;
            foreach (string patternName in LoadPatternNames)
            {
                values[0, columnIndex++] = "LC | " + patternName;
            }

            foreach (string comboName in LoadCombinationReferenceNames)
            {
                values[0, columnIndex++] = "COMBO | " + comboName;
            }

            for (int rowIndex = 0; rowIndex < Rows.Count; rowIndex++)
            {
                var row = Rows[rowIndex];
                values[rowIndex + 1, 0] = NormalizeName(row.LoadCombinationName) ?? string.Empty;
                values[rowIndex + 1, 1] = GetCombinationTypeDisplayName(row.CombinationType);

                columnIndex = 2;
                foreach (string patternName in LoadPatternNames)
                {
                    var factorResult = CreateExportFactor(row.GetLoadCaseFactor(patternName), rowIndex + 1, "LC | " + patternName);
                    if (!factorResult.IsSuccess)
                    {
                        return OperationResult<object[,]>.Failure(factorResult.Message);
                    }

                    values[rowIndex + 1, columnIndex++] = factorResult.Data;
                }

                foreach (string comboName in LoadCombinationReferenceNames)
                {
                    var factorResult = CreateExportFactor(row.GetLoadCombinationFactor(comboName), rowIndex + 1, "COMBO | " + comboName);
                    if (!factorResult.IsSuccess)
                    {
                        return OperationResult<object[,]>.Failure(factorResult.Message);
                    }

                    values[rowIndex + 1, columnIndex++] = factorResult.Data;
                }
            }

            return OperationResult<object[,]>.Success(values);
        }

        private static OperationResult<object> CreateExportFactor(string factorText, int rowNumber, string columnName)
        {
            if (string.IsNullOrWhiteSpace(factorText))
            {
                return OperationResult<object>.Success(string.Empty);
            }

            if (!TryParseDouble(factorText, out double factor))
            {
                return OperationResult<object>.Failure($"Row {rowNumber}, column {columnName}: value '{factorText}' is not numeric.");
            }

            return OperationResult<object>.Success(factor);
        }

        private static bool TryParseDouble(string value, out double result)
        {
            bool parsed = double.TryParse(value, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.CurrentCulture, out result)
                || double.TryParse(value, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out result);

            return parsed && !double.IsNaN(result) && !double.IsInfinity(result);
        }

        private IReadOnlyList<string> CreateDeletedLoadCombinationNames(IReadOnlyList<LoadCombinationMatrixRowDto> savedRows)
        {
            var currentNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            if (savedRows != null)
            {
                foreach (var row in savedRows)
                {
                    string name = row == null ? null : NormalizeName(row.LoadCombinationName);
                    if (!string.IsNullOrWhiteSpace(name))
                    {
                        currentNames.Add(name);
                    }
                }
            }

            var deletedNames = new List<string>();
            foreach (string originalName in _originalLoadCombinationNames)
            {
                if (!currentNames.Contains(originalName))
                {
                    deletedNames.Add(originalName);
                }
            }

            return deletedNames;
        }

        private static string NormalizeName(string name)
        {
            return string.IsNullOrWhiteSpace(name) ? null : name.Trim();
        }

        private static ObservableCollection<LoadCombinationTypeOption> CreateCombinationTypeOptions()
        {
            return new ObservableCollection<LoadCombinationTypeOption>
            {
                new LoadCombinationTypeOption { Value = (int)LoadCombinationType.LinearAdditive, DisplayName = "Linear Additive" },
                new LoadCombinationTypeOption { Value = (int)LoadCombinationType.Envelope, DisplayName = "Envelope" },
                new LoadCombinationTypeOption { Value = (int)LoadCombinationType.AbsoluteAdditive, DisplayName = "Absolute Additive" },
                new LoadCombinationTypeOption { Value = (int)LoadCombinationType.SRSS, DisplayName = "SRSS" },
                new LoadCombinationTypeOption { Value = (int)LoadCombinationType.RangeAdditive, DisplayName = "Range Additive" }
            };
        }

        private static bool IsSupportedCombinationType(int combinationType)
        {
            return combinationType == (int)LoadCombinationType.LinearAdditive
                || combinationType == (int)LoadCombinationType.Envelope
                || combinationType == (int)LoadCombinationType.AbsoluteAdditive
                || combinationType == (int)LoadCombinationType.SRSS
                || combinationType == (int)LoadCombinationType.RangeAdditive;
        }

        private string GetCombinationTypeDisplayName(int combinationType)
        {
            var option = CombinationTypeOptions.FirstOrDefault(x => x.Value == combinationType);
            return option == null
                ? combinationType.ToString(CultureInfo.InvariantCulture)
                : option.DisplayName;
        }
    }
}
