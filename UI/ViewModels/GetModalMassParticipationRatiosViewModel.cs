using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Input;
using ExcelCSIToolBox.Application.UseCases;
using ExcelCSIToolBox.Core.Abstractions.CSI;
using ExcelCSIToolBox.Core.Abstractions.Excel;
using ExcelCSIToolBox.Core.Common.Commands;
using ExcelCSIToolBox.Core.Common.Results;
using ExcelCSIToolBox.Data.DTOs.CSI;
using ExcelCSIToolBox.Infrastructure.Excel;
using ExcelCSIToolBoxAddIn.UI.Helpers;
using ExcelRange = Microsoft.Office.Interop.Excel.Range;

namespace ExcelCSIToolBoxAddIn.UI.ViewModels
{
    public class GetModalMassParticipationRatiosViewModel : ViewModelBase
    {
        private const string WorkbookStateKey = "ModalMassParticipationRatios";
        private readonly ICSISapModelConnectionService _csiConnectionService;
        private readonly IExcelOutputService _excelOutputService;
        private readonly GetModalMassParticipationRatiosUseCase _useCase;
        private string _anchorCellAddress;
        private string _statusText;
        private bool _isBusy;
        private bool _addHeaders;
        private bool _isUseActiveCellMode = true;
        private bool _isPickCellMode;
        private ExcelRange _pickedAnchorCell;
        private int _selectedLoadCaseCount;
        private IReadOnlyList<string> _selectedLoadCaseNames = new string[0];
        private PostprocessingWorkbookState _workbookState;
        private bool _isWorkbookStateLoaded;
        private string _etabsModelName = "ETABS Model: Not attached";

        public GetModalMassParticipationRatiosViewModel(
            CsiToolboxUseCaseBundle useCases,
            ICSISapModelConnectionService csiConnectionService,
            IExcelOutputService excelOutputService)
        {
            if (useCases == null) throw new ArgumentNullException(nameof(useCases));
            _csiConnectionService = csiConnectionService ?? throw new ArgumentNullException(nameof(csiConnectionService));
            _excelOutputService = excelOutputService ?? throw new ArgumentNullException(nameof(excelOutputService));
            _useCase = useCases.GetModalMassParticipationRatios ?? throw new ArgumentNullException(nameof(useCases.GetModalMassParticipationRatios));

            ModalLoadCases = new ObservableCollection<BaseReactionOutputCaseViewModel>();
            _workbookState = PostprocessingWorkbookStateStore.Load(WorkbookStateKey);
            RestoreWorkbookState();
            _isWorkbookStateLoaded = true;
            LoadOutputCasesCommand = new RelayCommand(LoadOutputCases, () => !IsBusy);
            PickAnchorCellCommand = new RelayCommand(() => { IsPickCellMode = true; }, () => !IsBusy);
            RunCommand = new RelayCommand(Run, () => !IsBusy);
            CancelCommand = new RelayCommand(() => RequestClose?.Invoke(this, EventArgs.Empty));

            RefreshAnchorDisplay();
            LoadOutputCases();
        }

        public event EventHandler RequestClose;

        public ObservableCollection<BaseReactionOutputCaseViewModel> ModalLoadCases { get; private set; }

        public ICommand LoadOutputCasesCommand { get; private set; }
        public ICommand PickAnchorCellCommand { get; private set; }
        public ICommand RunCommand { get; private set; }
        public ICommand CancelCommand { get; private set; }

        public string AnchorCellAddress
        {
            get { return _anchorCellAddress; }
            private set
            {
                _anchorCellAddress = value;
                OnPropertyChanged();
            }
        }

        public string StatusText
        {
            get { return _statusText; }
            private set
            {
                _statusText = value;
                OnPropertyChanged();
            }
        }

        public string EtabsModelName
        {
            get { return _etabsModelName; }
            private set
            {
                _etabsModelName = value;
                OnPropertyChanged();
            }
        }

        public bool IsBusy
        {
            get { return _isBusy; }
            private set
            {
                _isBusy = value;
                OnPropertyChanged();
                RaiseCommandState(LoadOutputCasesCommand);
                RaiseCommandState(PickAnchorCellCommand);
                RaiseCommandState(RunCommand);
            }
        }

        public bool AddHeaders
        {
            get { return _addHeaders; }
            set
            {
                _addHeaders = value;
                OnPropertyChanged();
                SaveWorkbookState();
            }
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
                    _pickedAnchorCell = null;
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
            get
            {
                return IsPickCellMode
                    ? "Picked cell is fixed until changed."
                    : "Uses the current Excel active cell when Run is clicked.";
            }
        }

        public int SelectedLoadCaseCount
        {
            get { return _selectedLoadCaseCount; }
            private set
            {
                _selectedLoadCaseCount = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(LoadCaseSelectionText));
            }
        }

        public string LoadCaseSelectionText
        {
            get { return $"{SelectedLoadCaseCount}/{ModalLoadCases.Count} selected"; }
        }

        public void UpdateSelectionCount(int selectedLoadCaseCount)
        {
            SelectedLoadCaseCount = selectedLoadCaseCount;
        }

        public void UpdateSelectedOutputCases(System.Collections.IList selectedLoadCases)
        {
            _selectedLoadCaseNames = GetSelectedOutputCaseNames(selectedLoadCases);
            UpdateSelectionCount(_selectedLoadCaseNames.Count);
            SaveWorkbookState();
        }

        public void RestoreSavedSelections(System.Collections.IList selectedLoadCases)
        {
            RestoreSelectedItems(selectedLoadCases, ModalLoadCases, _workbookState.LoadCaseNames);
            UpdateSelectedOutputCases(selectedLoadCases);
        }

        public void RefreshAnchorDisplay()
        {
            if (IsUseActiveCellMode)
            {
                RefreshActiveCellDisplay();
            }
        }

        private void LoadOutputCases()
        {
            if (!EnsureEtabs())
            {
                return;
            }

            try
            {
                IsBusy = true;
                StatusText = "Loading ETABS modal load cases...";
                var result = _csiConnectionService.GetAnalysisOutputCases();
                if (!result.IsSuccess)
                {
                    StatusText = result.Message;
                    ShowWarning(result.Message);
                    return;
                }

                ModalLoadCases.Clear();
                if (result.Data != null)
                {
                    foreach (var outputCase in result.Data)
                    {
                        if (!outputCase.IsLoadCombination &&
                            string.Equals(outputCase.Type, "Modal", StringComparison.OrdinalIgnoreCase))
                        {
                            ModalLoadCases.Add(new BaseReactionOutputCaseViewModel(outputCase));
                        }
                    }
                }

                UpdateSelectionCount(0);
                OnPropertyChanged(nameof(LoadCaseSelectionText));
                StatusText = ModalLoadCases.Count == 0
                    ? "No ETABS modal load cases were found."
                    : $"Loaded {ModalLoadCases.Count} modal load case(s).";
            }
            catch (Exception ex)
            {
                StatusText = "Failed to load ETABS modal load cases.";
                ShowError($"Failed to load ETABS modal load cases: {ex.Message}");
            }
            finally
            {
                IsBusy = false;
            }
        }

        public void Run(System.Collections.IList selectedLoadCases)
        {
            if (!EnsureEtabs() || !PrepareAnchorCellForWrite())
            {
                return;
            }

            try
            {
                RaiseRequestHide();

                var selectedCases = GetSelectedOutputCases(selectedLoadCases);
                if (selectedCases.Count == 0)
                {
                    ShowWarning("Select at least one ETABS modal load case.");
                    return;
                }

                try
                {
                    IsBusy = true;
                    StatusText = "Extracting ETABS Modal Mass Participation Ratios...";
                    var result = _useCase.Execute(selectedCases);
                    if (!result.IsSuccess)
                    {
                        StatusText = result.Message;
                        ShowWarning(result.Message);
                        return;
                    }

                    if (result.Data == null || result.Data.Count == 0)
                    {
                        StatusText = "ETABS returned no Modal Mass Participation Ratio records.";
                        MessageBox.Show(
                            "ETABS returned no Modal Mass Participation Ratio records for the selected modal load cases. Nothing was written to Excel.",
                            "Modal Mass Participation Ratios",
                            MessageBoxButton.OK,
                            MessageBoxImage.Information);
                        return;
                    }

                    object[,] values = CreateOutputValues(result.Data, AddHeaders);
                    OperationResult writeResult = _excelOutputService.WriteValuesToActiveCell(
                        values,
                        $"Successfully wrote {result.Data.Count} Modal Mass Participation Ratio record(s) to Excel.",
                        AddHeaders);

                    StatusText = writeResult.Message;
                    MessageBox.Show(
                        writeResult.Message,
                        "Modal Mass Participation Ratios",
                        MessageBoxButton.OK,
                        writeResult.IsSuccess ? MessageBoxImage.Information : MessageBoxImage.Warning);
                }
                catch (Exception ex)
                {
                    StatusText = "Failed to extract Modal Mass Participation Ratios.";
                    ShowError($"Failed to extract Modal Mass Participation Ratios: {ex.Message}");
                }
                finally
                {
                    IsBusy = false;
                }
            }
            finally
            {
                RaiseRequestShow();
            }
        }

        private void Run()
        {
            Run(null);
        }

        private static List<CSISapModelOutputCaseDTO> GetSelectedOutputCases(System.Collections.IList selectedItems)
        {
            var selectedCases = new List<CSISapModelOutputCaseDTO>();
            if (selectedItems == null)
            {
                return selectedCases;
            }

            foreach (var selectedItem in selectedItems)
            {
                var item = selectedItem as BaseReactionOutputCaseViewModel;
                if (item != null && item.OutputCase != null)
                {
                    selectedCases.Add(item.OutputCase);
                }
            }

            return selectedCases;
        }

        private static object[,] CreateOutputValues(IReadOnlyList<CSISapModelModalMassParticipationRowDTO> rows, bool addHeaders)
        {
            int headerOffset = addHeaders ? 1 : 0;
            var values = new object[rows.Count + headerOffset, 15];

            if (addHeaders)
            {
                string[] headers = new[]
                {
                    "Output Case",
                    "Step Number",
                    "Period",
                    "UX",
                    "UY",
                    "UZ",
                    "Sum UX",
                    "Sum UY",
                    "Sum UZ",
                    "RX",
                    "RY",
                    "RZ",
                    "Sum RX",
                    "Sum RY",
                    "Sum RZ"
                };

                for (int col = 0; col < headers.Length; col++)
                {
                    values[0, col] = headers[col];
                }
            }

            for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++)
            {
                CSISapModelModalMassParticipationRowDTO row = rows[rowIndex];
                int targetRowIndex = rowIndex + headerOffset;
                values[targetRowIndex, 0] = row.OutputCase;
                values[targetRowIndex, 1] = row.StepNumber;
                values[targetRowIndex, 2] = row.Period;
                values[targetRowIndex, 3] = row.UX;
                values[targetRowIndex, 4] = row.UY;
                values[targetRowIndex, 5] = row.UZ;
                values[targetRowIndex, 6] = row.SumUX;
                values[targetRowIndex, 7] = row.SumUY;
                values[targetRowIndex, 8] = row.SumUZ;
                values[targetRowIndex, 9] = row.RX;
                values[targetRowIndex, 10] = row.RY;
                values[targetRowIndex, 11] = row.RZ;
                values[targetRowIndex, 12] = row.SumRX;
                values[targetRowIndex, 13] = row.SumRY;
                values[targetRowIndex, 14] = row.SumRZ;
            }

            return values;
        }

        private bool PrepareAnchorCellForWrite()
        {
            if (IsPickCellMode)
            {
                if (_pickedAnchorCell == null && !PickAnchorCell())
                {
                    return false;
                }

                try
                {
                    _pickedAnchorCell.Select();
                    AnchorCellAddress = FormatAddress(_pickedAnchorCell);
                    return true;
                }
                catch (Exception ex)
                {
                    ShowWarning($"Failed to activate the picked Excel anchor cell: {ex.Message}");
                    return false;
                }
            }

            ExcelRange activeCell = GetActiveExcelCell();
            if (activeCell == null)
            {
                AnchorCellAddress = string.Empty;
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
                if (excelApp == null)
                {
                    ShowWarning("Excel application is not available.");
                    return false;
                }

                object result = excelApp.InputBox(
                    AddHeaders
                        ? "Select the top-left anchor cell where Modal Mass Participation Ratio headers should start. Data will start one row below."
                        : "Select the top-left anchor cell where the first Modal Mass Participation Ratio data row should start. Headers are excluded.",
                    "Modal Mass Participation Ratios",
                    Type: 8);

                if (result is bool && (bool)result == false)
                {
                    return false;
                }

                var selectedRange = result as ExcelRange;
                ExcelRange startCell = selectedRange == null ? null : selectedRange.Cells[1, 1] as ExcelRange;
                if (startCell == null)
                {
                    ShowWarning("No Excel anchor cell was selected.");
                    return false;
                }

                _pickedAnchorCell = startCell;
                startCell.Select();
                AnchorCellAddress = FormatAddress(startCell);
                StatusText = $"Anchor cell set to {AnchorCellAddress}.";
                SaveWorkbookState();
                return true;
            }
            catch (Exception ex)
            {
                ShowError($"Failed to select the Excel anchor cell: {ex.Message}");
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
            if (activeCell == null)
            {
                AnchorCellAddress = string.Empty;
                StatusText = "Select an Excel anchor cell for output.";
                return;
            }

            AnchorCellAddress = FormatAddress(activeCell);
            StatusText = $"Anchor cell set to {AnchorCellAddress}.";
        }

        private bool EnsureEtabs()
        {
            if (!string.Equals(_csiConnectionService.ProductName, "ETABS", StringComparison.OrdinalIgnoreCase))
            {
                ShowWarning("Modal Mass Participation Ratios is available from the ETABS Toolbox only.");
                return false;
            }

            var connectionResult = _csiConnectionService.GetCurrentConnection();
            if (connectionResult.IsSuccess)
            {
                UpdateEtabsModelName(connectionResult.Data);
                return true;
            }

            var attachResult = _csiConnectionService.TryAttachToRunningInstance();
            if (attachResult.IsSuccess)
            {
                UpdateEtabsModelName(attachResult.Data);
                return true;
            }

            ShowWarning(string.IsNullOrWhiteSpace(attachResult.Message)
                ? "No ETABS model is currently connected. Please attach to a running ETABS instance."
                : attachResult.Message);
            return false;
        }

        private static ExcelRange GetActiveExcelCell()
        {
            try
            {
                var excelApp = ExcelApplicationProvider.GetApplication();
                if (excelApp == null)
                {
                    return null;
                }

                var selectedRange = excelApp.Selection as ExcelRange;
                if (selectedRange != null)
                {
                    return selectedRange.Cells[1, 1] as ExcelRange;
                }

                return excelApp.ActiveCell as ExcelRange;
            }
            catch
            {
                return null;
            }
        }

        private static string FormatAddress(ExcelRange cell)
        {
            if (cell == null)
            {
                return string.Empty;
            }

            string address = cell.Address[RowAbsolute: false, ColumnAbsolute: false];
            string sheetName = cell.Worksheet == null ? string.Empty : cell.Worksheet.Name;
            return string.IsNullOrWhiteSpace(sheetName) ? address : $"{sheetName}!{address}";
        }

        private void RestoreWorkbookState()
        {
            if (_workbookState == null)
            {
                return;
            }

            AddHeaders = _workbookState.AddHeaders;
            if (_workbookState.UsePickedAnchor)
            {
                ExcelRange anchorCell = PostprocessingWorkbookStateStore.TryGetAnchorCell(_workbookState.AnchorAddress);
                if (anchorCell != null)
                {
                    _pickedAnchorCell = anchorCell;
                    _isUseActiveCellMode = false;
                    _isPickCellMode = true;
                    AnchorCellAddress = FormatAddress(anchorCell);
                    OnPropertyChanged(nameof(IsUseActiveCellMode));
                    OnPropertyChanged(nameof(IsPickCellMode));
                    OnPropertyChanged(nameof(AnchorModeText));
                }
            }
        }

        private void SaveWorkbookState()
        {
            if (!_isWorkbookStateLoaded)
            {
                return;
            }

            PostprocessingWorkbookStateStore.Save(WorkbookStateKey, new PostprocessingWorkbookState
            {
                AddHeaders = AddHeaders,
                UsePickedAnchor = IsPickCellMode,
                AnchorAddress = IsPickCellMode ? AnchorCellAddress : string.Empty,
                LoadCaseNames = _selectedLoadCaseNames,
                LoadCombinationNames = new string[0]
            });
        }

        private void UpdateEtabsModelName(CSISapModelConnectionInfoDTO connection)
        {
            string modelName = connection == null ? string.Empty : connection.ModelFileName;
            EtabsModelName = string.IsNullOrWhiteSpace(modelName) ? "ETABS Model: Untitled" : $"ETABS Model: {modelName}";
        }

        private static IReadOnlyList<string> GetSelectedOutputCaseNames(System.Collections.IList selectedItems)
        {
            var names = new List<string>();
            if (selectedItems == null)
            {
                return names;
            }

            foreach (object selectedItem in selectedItems)
            {
                var item = selectedItem as BaseReactionOutputCaseViewModel;
                if (item != null && !string.IsNullOrWhiteSpace(item.Name))
                {
                    names.Add(item.Name);
                }
            }

            return names;
        }

        private static void RestoreSelectedItems(
            System.Collections.IList selectedItems,
            IEnumerable<BaseReactionOutputCaseViewModel> availableItems,
            IReadOnlyList<string> selectedNames)
        {
            if (selectedItems == null || availableItems == null || selectedNames == null)
            {
                return;
            }

            selectedItems.Clear();
            var nameSet = new HashSet<string>(selectedNames, StringComparer.OrdinalIgnoreCase);
            foreach (BaseReactionOutputCaseViewModel item in availableItems)
            {
                if (item != null && nameSet.Contains(item.Name))
                {
                    selectedItems.Add(item);
                }
            }
        }

        private static void RaiseCommandState(ICommand command)
        {
            var relayCommand = command as IRelayCommand;
            if (relayCommand != null)
            {
                relayCommand.RaiseCanExecuteChanged();
            }
        }

        private static void ShowWarning(string message)
        {
            MessageBox.Show(
                string.IsNullOrWhiteSpace(message) ? "The operation could not be completed." : message,
                "Modal Mass Participation Ratios",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
        }

        private static void ShowError(string message)
        {
            MessageBox.Show(
                string.IsNullOrWhiteSpace(message) ? "An unexpected error occurred." : message,
                "Modal Mass Participation Ratios",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
        }
    }
}
