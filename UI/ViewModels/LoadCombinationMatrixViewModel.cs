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
using ExcelCSIToolBox.Data.DTOs.CSI;

namespace ExcelCSIToolBoxAddIn.UI.ViewModels
{
    public class LoadCombinationMatrixViewModel : ViewModelBase
    {
        private readonly string _productTitle;
        private IReadOnlyList<LoadCombinationMatrixRowDto> _savedRows = new List<LoadCombinationMatrixRowDto>();

        public LoadCombinationMatrixViewModel(
            LoadCombinationMatrixDto initialMatrix,
            string productTitle)
        {
            _productTitle = string.IsNullOrWhiteSpace(productTitle) ? "CSI Toolbox" : productTitle;

            Rows = new ObservableCollection<LoadCombinationMatrixRowViewModel>();
            LoadPatternNames = new ObservableCollection<string>();

            AddRowCommand = new RelayCommand(AddRow);
            DeleteSelectedRowsCommand = new RelayCommand<IList>(DeleteSelectedRows);
            SaveCommand = new RelayCommand(Save);
            CancelCommand = new RelayCommand(Cancel);

            LoadMatrix(initialMatrix ?? new LoadCombinationMatrixDto());
        }

        public ObservableCollection<LoadCombinationMatrixRowViewModel> Rows { get; }
        public ObservableCollection<string> LoadPatternNames { get; }

        public ICommand AddRowCommand { get; }
        public ICommand DeleteSelectedRowsCommand { get; }
        public ICommand SaveCommand { get; }
        public ICommand CancelCommand { get; }

        public bool WasSaved { get; private set; }

        public IReadOnlyList<LoadCombinationMatrixRowDto> SavedRows
        {
            get { return _savedRows; }
        }

        public event EventHandler RequestClose;

        private void LoadMatrix(LoadCombinationMatrixDto matrix)
        {
            LoadPatternNames.Clear();
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

            if (matrix.Rows != null)
            {
                foreach (var row in matrix.Rows)
                {
                    Rows.Add(LoadCombinationMatrixRowViewModel.FromDto(row, LoadPatternNames));
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

        private void Save()
        {
            var validation = ValidateAndCreateDtos();
            if (!validation.IsSuccess)
            {
                MessageBox.Show(validation.Message, _productTitle, MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            _savedRows = validation.Data;
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

            int rowNumber = 1;
            foreach (var row in Rows)
            {
                string name = string.IsNullOrWhiteSpace(row.LoadCombinationName) ? null : row.LoadCombinationName.Trim();
                if (string.IsNullOrWhiteSpace(name))
                {
                    errors.Add($"Row {rowNumber}: LoadCombinationName is required.");
                }
                else if (!names.Add(name))
                {
                    errors.Add($"Row {rowNumber}: duplicate LoadCombinationName '{name}'.");
                }

                var dto = new LoadCombinationMatrixRowDto
                {
                    LoadCombinationName = name,
                    CombinationType = row.CombinationType
                };

                foreach (string patternName in LoadPatternNames)
                {
                    string factorText = row[patternName];
                    if (string.IsNullOrWhiteSpace(factorText))
                    {
                        continue;
                    }

                    if (!TryParseDouble(factorText, out double factor))
                    {
                        errors.Add($"Row {rowNumber}, column {patternName}: value '{factorText}' is not numeric.");
                        continue;
                    }

                    dto.Factors[patternName] = factor;
                }

                dtos.Add(dto);
                rowNumber++;
            }

            if (dtos.Count == 0)
            {
                errors.Add("No load combination rows are available to save.");
            }

            if (errors.Count > 0)
            {
                return OperationResult<IReadOnlyList<LoadCombinationMatrixRowDto>>.Failure("Matrix validation failed: " + string.Join(" ", errors));
            }

            return OperationResult<IReadOnlyList<LoadCombinationMatrixRowDto>>.Success(dtos);
        }

        private static bool TryParseDouble(string value, out double result)
        {
            bool parsed = double.TryParse(value, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.CurrentCulture, out result)
                || double.TryParse(value, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out result);

            return parsed && !double.IsNaN(result) && !double.IsInfinity(result);
        }
    }
}
