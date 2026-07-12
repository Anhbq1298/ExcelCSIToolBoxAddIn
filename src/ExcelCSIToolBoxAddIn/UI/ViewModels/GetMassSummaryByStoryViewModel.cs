using System;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Windows;
using System.Windows.Input;
using ExcelCSIToolBox.Application.Composition;
using ExcelCSIToolBox.Application.Features.AnalysisResults;
using ExcelCSIToolBox.Core.Abstractions.CSI;
using ExcelCSIToolBox.Core.Abstractions.Excel;
using ExcelCSIToolBoxAddIn.UI.Common.Commands;
using ExcelCSIToolBox.Core.Common.Results;
using ExcelCSIToolBox.Core.Contracts.CSI;
using ExcelCSIToolBox.Infrastructure.Excel.Interop;
using ExcelCSIToolBoxAddIn.UI.Helpers;
using ExcelRange = Microsoft.Office.Interop.Excel.Range;

namespace ExcelCSIToolBoxAddIn.UI.ViewModels
{
    public class GetMassSummaryByStoryViewModel : ViewModelBase
    {
        private const string WorkbookStateKey = "MassSummaryByStory";
        private readonly ICSISapModelConnectionService _csiConnectionService;
        private readonly IExcelOutputService _excelOutputService;
        private readonly GetMassSummaryByStoryUseCase _useCase;
        private string _anchorCellAddress;
        private string _statusText;
        private string _etabsModelName = "ETABS Model: Not attached";
        private bool _isBusy;
        private bool _addHeaders;
        private bool _isUseActiveCellMode = true;
        private bool _isPickCellMode;
        private bool _isWorkbookStateLoaded;
        private ExcelRange _pickedAnchorCell;
        private MassUnitOption _selectedMassUnitOption;

        public GetMassSummaryByStoryViewModel(
            CsiToolboxUseCaseBundle useCases,
            ICSISapModelConnectionService csiConnectionService,
            IExcelOutputService excelOutputService)
        {
            _csiConnectionService = csiConnectionService ?? throw new ArgumentNullException(nameof(csiConnectionService));
            _excelOutputService = excelOutputService ?? throw new ArgumentNullException(nameof(excelOutputService));
            _useCase = (useCases ?? throw new ArgumentNullException(nameof(useCases))).GetMassSummaryByStory;

            MassUnitOptions = new ObservableCollection<MassUnitOption>
            {
                new MassUnitOption("kg", 9.80665d),
                new MassUnitOption("ton", 9.80665d / 1000d),
                new MassUnitOption("kN", 9.80665d * 9.80665d / 1000d)
            };
            SelectedMassUnitOption = MassUnitOptions[2];
            RestoreWorkbookState();
            _isWorkbookStateLoaded = true;
            PickAnchorCellCommand = new RelayCommand(() => { IsPickCellMode = true; }, () => !IsBusy);
            RunCommand = new RelayCommand(Run, () => !IsBusy);
            CancelCommand = new RelayCommand(() => RequestClose?.Invoke(this, EventArgs.Empty));

            RefreshAnchorDisplay();
            EnsureEtabs();
        }

        public event EventHandler RequestClose;
        public ICommand PickAnchorCellCommand { get; private set; }
        public ICommand RunCommand { get; private set; }
        public ICommand CancelCommand { get; private set; }

        public ObservableCollection<MassUnitOption> MassUnitOptions { get; private set; }

        public MassUnitOption SelectedMassUnitOption
        {
            get { return _selectedMassUnitOption; }
            set { _selectedMassUnitOption = value; OnPropertyChanged(); SaveWorkbookState(); }
        }

        public string EtabsModelName
        {
            get { return _etabsModelName; }
            private set { _etabsModelName = value; OnPropertyChanged(); }
        }

        public string AnchorCellAddress
        {
            get { return _anchorCellAddress; }
            private set { _anchorCellAddress = value; OnPropertyChanged(); }
        }

        public string StatusText
        {
            get { return _statusText; }
            private set { _statusText = value; OnPropertyChanged(); }
        }

        public bool IsBusy
        {
            get { return _isBusy; }
            private set
            {
                _isBusy = value;
                OnPropertyChanged();
                RaiseCommandState(PickAnchorCellCommand);
                RaiseCommandState(RunCommand);
            }
        }

        public bool AddHeaders
        {
            get { return _addHeaders; }
            set { _addHeaders = value; OnPropertyChanged(); SaveWorkbookState(); }
        }

        public bool IsUseActiveCellMode
        {
            get { return _isUseActiveCellMode; }
            set
            {
                if (!value || _isUseActiveCellMode == value) return;
                _isUseActiveCellMode = true;
                _isPickCellMode = false;
                _pickedAnchorCell = null;
                OnPropertyChanged();
                OnPropertyChanged(nameof(IsPickCellMode));
                OnPropertyChanged(nameof(AnchorModeText));
                RefreshActiveCellDisplay();
                SaveWorkbookState();
            }
        }

        public bool IsPickCellMode
        {
            get { return _isPickCellMode; }
            set
            {
                if (!value || _isPickCellMode == value) return;
                _isPickCellMode = true;
                _isUseActiveCellMode = false;
                OnPropertyChanged();
                OnPropertyChanged(nameof(IsUseActiveCellMode));
                OnPropertyChanged(nameof(AnchorModeText));
                if (!PickAnchorCell())
                {
                    _isPickCellMode = false;
                    _isUseActiveCellMode = true;
                    OnPropertyChanged();
                    OnPropertyChanged(nameof(IsUseActiveCellMode));
                    OnPropertyChanged(nameof(AnchorModeText));
                    RefreshActiveCellDisplay();
                }

                SaveWorkbookState();
            }
        }

        public string AnchorModeText
        {
            get { return IsPickCellMode ? "Picked cell is fixed until changed." : "Uses the current Excel active cell when Run is clicked."; }
        }

        public void RefreshAnchorDisplay()
        {
            if (IsUseActiveCellMode) RefreshActiveCellDisplay();
        }

        private void Run()
        {
            if (!EnsureEtabs() || !PrepareAnchorCellForWrite()) return;

            try
            {
                RaiseRequestHide();
                IsBusy = true;
                StatusText = "Extracting ETABS Mass Summary by Story...";
                var result = _useCase.Execute();
                if (!result.IsSuccess)
                {
                    StatusText = result.Message;
                    ShowWarning(result.Message);
                    return;
                }

                if (result.Data == null || result.Data.Rows == null || result.Data.Rows.Count == 0)
                {
                    StatusText = "ETABS returned no Mass Summary by Story records.";
                    ShowInformation("ETABS returned no Mass Summary by Story records. Nothing was written to Excel.");
                    return;
                }

                object[,] values = CreateOutputValues(result.Data, AddHeaders, SelectedMassUnitOption);
                OperationResult writeResult = _excelOutputService.WriteValuesToActiveCell(
                    values,
                    $"Successfully wrote {result.Data.Rows.Count} Mass Summary by Story record(s) to Excel.",
                    AddHeaders);
                StatusText = writeResult.Message;
                ShowInformation(writeResult.Message, writeResult.IsSuccess ? MessageBoxImage.Information : MessageBoxImage.Warning);
            }
            catch (Exception ex)
            {
                StatusText = "Failed to extract Mass Summary by Story.";
                ShowInformation($"Failed to extract Mass Summary by Story: {ex.Message}", MessageBoxImage.Error);
            }
            finally
            {
                IsBusy = false;
                RaiseRequestShow();
            }
        }

        private bool EnsureEtabs()
        {
            var connectionResult = _csiConnectionService.GetCurrentConnection();
            if (!connectionResult.IsSuccess)
            {
                connectionResult = _csiConnectionService.TryAttachToRunningInstance();
            }

            if (connectionResult.IsSuccess)
            {
                string modelName = connectionResult.Data == null ? string.Empty : connectionResult.Data.ModelFileName;
                EtabsModelName = string.IsNullOrWhiteSpace(modelName) ? "ETABS Model: Untitled" : $"ETABS Model: {modelName}";
                return true;
            }

            ShowWarning(string.IsNullOrWhiteSpace(connectionResult.Message)
                ? "No ETABS model is currently connected. Please attach to a running ETABS instance."
                : connectionResult.Message);
            return false;
        }

        private bool PrepareAnchorCellForWrite()
        {
            if (IsPickCellMode)
            {
                if (_pickedAnchorCell == null && !PickAnchorCell()) return false;
                _pickedAnchorCell.Select();
                AnchorCellAddress = FormatAddress(_pickedAnchorCell);
                return true;
            }

            ExcelRange activeCell = GetActiveExcelCell();
            if (activeCell == null)
            {
                ShowWarning("Select an Excel active cell before running.");
                return false;
            }

            AnchorCellAddress = FormatAddress(activeCell);
            return true;
        }

        private bool PickAnchorCell()
        {
            try
            {
                RaiseRequestHide();
                var excelApp = ExcelApplicationProvider.GetApplication();
                object result = excelApp == null ? null : excelApp.InputBox(
                    AddHeaders
                        ? "Select the top-left anchor cell where Mass Summary by Story headers should start. Data will start one row below."
                        : "Select the top-left anchor cell where the first Mass Summary by Story data row should start. Headers are excluded.",
                    "Mass Summary by Story", Type: 8);
                if (result is bool && (bool)result == false) return false;

                var selectedRange = result as ExcelRange;
                ExcelRange startCell = selectedRange == null ? null : selectedRange.Cells[1, 1] as ExcelRange;
                if (startCell == null) return false;

                _pickedAnchorCell = startCell;
                startCell.Select();
                AnchorCellAddress = FormatAddress(startCell);
                StatusText = $"Anchor cell set to {AnchorCellAddress}.";
                SaveWorkbookState();
                return true;
            }
            catch (Exception ex)
            {
                ShowInformation($"Failed to select the Excel anchor cell: {ex.Message}", MessageBoxImage.Error);
                return false;
            }
            finally
            {
                RaiseRequestShow();
            }
        }

        private void RefreshActiveCellDisplay()
        {
            ExcelRange activeCell = GetActiveExcelCell();
            AnchorCellAddress = activeCell == null ? string.Empty : FormatAddress(activeCell);
        }

        private void RestoreWorkbookState()
        {
            PostprocessingWorkbookState state = PostprocessingWorkbookStateStore.Load(WorkbookStateKey);
            foreach (MassUnitOption unitOption in MassUnitOptions)
            {
                if (string.Equals(unitOption.Label, state.UnitLabel, StringComparison.OrdinalIgnoreCase))
                {
                    SelectedMassUnitOption = unitOption;
                    break;
                }
            }

            AddHeaders = state.AddHeaders;
            if (!state.UsePickedAnchor) return;

            ExcelRange anchorCell = PostprocessingWorkbookStateStore.TryGetAnchorCell(state.AnchorAddress);
            if (anchorCell == null) return;

            _pickedAnchorCell = anchorCell;
            _isUseActiveCellMode = false;
            _isPickCellMode = true;
            AnchorCellAddress = FormatAddress(anchorCell);
            OnPropertyChanged(nameof(IsUseActiveCellMode));
            OnPropertyChanged(nameof(IsPickCellMode));
            OnPropertyChanged(nameof(AnchorModeText));
        }

        private void SaveWorkbookState()
        {
            if (!_isWorkbookStateLoaded) return;
            PostprocessingWorkbookStateStore.Save(WorkbookStateKey, new PostprocessingWorkbookState
            {
                UnitLabel = SelectedMassUnitOption == null ? string.Empty : SelectedMassUnitOption.Label,
                AddHeaders = AddHeaders,
                UsePickedAnchor = IsPickCellMode,
                AnchorAddress = IsPickCellMode ? AnchorCellAddress : string.Empty,
                LoadCaseNames = new string[0],
                LoadCombinationNames = new string[0]
            });
        }

        private static object[,] CreateOutputValues(CSISapModelDisplayTableDTO table, bool addHeaders, MassUnitOption unitOption)
        {
            int fieldCount = table.FieldKeys == null ? 0 : table.FieldKeys.Count;
            int rowCount = table.Rows == null ? 0 : table.Rows.Count;
            int headerOffset = addHeaders ? 1 : 0;
            var values = new object[rowCount + headerOffset, fieldCount];
            if (addHeaders)
            {
                for (int fieldIndex = 0; fieldIndex < fieldCount; fieldIndex++)
                {
                    string fieldKey = table.FieldKeys[fieldIndex];
                    values[0, fieldIndex] = unitOption != null && !IsTextField(fieldKey)
                        ? $"{fieldKey} ({unitOption.Label})"
                        : fieldKey;
                }
            }

            for (int rowIndex = 0; rowIndex < rowCount; rowIndex++)
            {
                object[] row = table.Rows[rowIndex];
                for (int fieldIndex = 0; fieldIndex < fieldCount; fieldIndex++)
                {
                    object value = row != null && fieldIndex < row.Length ? row[fieldIndex] : string.Empty;
                    values[rowIndex + headerOffset, fieldIndex] = ScaleMassValue(table.FieldKeys[fieldIndex], value, unitOption);
                }
            }

            return values;
        }

        private static object ScaleMassValue(string fieldKey, object value, MassUnitOption unitOption)
        {
            if (unitOption == null || IsTextField(fieldKey) || value == null)
            {
                return value;
            }

            double number;
            string sourceValue = Convert.ToString(value, CultureInfo.InvariantCulture);
            if (double.TryParse(sourceValue, NumberStyles.Float, CultureInfo.InvariantCulture, out number) ||
                double.TryParse(sourceValue, NumberStyles.Float, CultureInfo.CurrentCulture, out number))
            {
                return number * unitOption.ScaleFactor;
            }

            return value;
        }

        private static bool IsTextField(string fieldKey)
        {
            return string.Equals(fieldKey, "Story", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(fieldKey, "Story Name", StringComparison.OrdinalIgnoreCase);
        }

        private static ExcelRange GetActiveExcelCell()
        {
            try
            {
                var excelApp = ExcelApplicationProvider.GetApplication();
                var selectedRange = excelApp == null ? null : excelApp.Selection as ExcelRange;
                return selectedRange == null ? excelApp == null ? null : excelApp.ActiveCell as ExcelRange : selectedRange.Cells[1, 1] as ExcelRange;
            }
            catch { return null; }
        }

        private static string FormatAddress(ExcelRange cell)
        {
            string address = cell.Address[RowAbsolute: false, ColumnAbsolute: false];
            string sheetName = cell.Worksheet == null ? string.Empty : cell.Worksheet.Name;
            return string.IsNullOrWhiteSpace(sheetName) ? address : $"{sheetName}!{address}";
        }

        private static void RaiseCommandState(ICommand command)
        {
            var relayCommand = command as IRelayCommand;
            if (relayCommand != null) relayCommand.RaiseCanExecuteChanged();
        }

        private static void ShowWarning(string message)
        {
            ShowInformation(message, MessageBoxImage.Warning);
        }

        private static void ShowInformation(string message, MessageBoxImage image = MessageBoxImage.Information)
        {
            MessageBox.Show(message, "Mass Summary by Story", MessageBoxButton.OK, image);
        }
    }

    public class MassUnitOption
    {
        public MassUnitOption(string label, double scaleFactor)
        {
            Label = label;
            ScaleFactor = scaleFactor;
        }

        public string Label { get; private set; }
        public double ScaleFactor { get; private set; }
    }
}
