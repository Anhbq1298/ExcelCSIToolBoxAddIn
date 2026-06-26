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
    public class GetBaseReactionsViewModel : ViewModelBase
    {
        private const string WorkbookStateKey = "BaseReactions";
        private readonly ICSISapModelConnectionService _csiConnectionService;
        private readonly IExcelOutputService _excelOutputService;
        private readonly GetBaseReactionsUseCase _getBaseReactionsUseCase;
        private string _anchorCellAddress;
        private string _statusText;
        private bool _isBusy;
        private bool _addHeaders;
        private bool _isUseActiveCellMode = true;
        private bool _isPickCellMode;
        private ExcelRange _pickedAnchorCell;
        private int _selectedLoadCaseCount;
        private int _selectedLoadCombinationCount;
        private IReadOnlyList<string> _selectedLoadCaseNames = new string[0];
        private IReadOnlyList<string> _selectedLoadCombinationNames = new string[0];
        private PostprocessingWorkbookState _workbookState;
        private bool _isWorkbookStateLoaded;
        private string _etabsModelName = "ETABS Model: Not attached";

        public GetBaseReactionsViewModel(
            CsiToolboxUseCaseBundle useCases,
            ICSISapModelConnectionService csiConnectionService,
            IExcelOutputService excelOutputService)
        {
            if (useCases == null) throw new ArgumentNullException(nameof(useCases));
            _csiConnectionService = csiConnectionService ?? throw new ArgumentNullException(nameof(csiConnectionService));
            _excelOutputService = excelOutputService ?? throw new ArgumentNullException(nameof(excelOutputService));
            _getBaseReactionsUseCase = useCases.GetBaseReactions ?? throw new ArgumentNullException(nameof(useCases.GetBaseReactions));

            UnitOptions = new ObservableCollection<BaseReactionUnitOption>
            {
                new BaseReactionUnitOption("N-mm", 9, "N", "N-mm", "mm"),
                new BaseReactionUnitOption("kN-m", 6, "kN", "kN-m", "m"),
                new BaseReactionUnitOption("kip-ft", 4, "kip", "kip-ft", "ft"),
                new BaseReactionUnitOption("lb-in", 1, "lb", "lb-in", "in")
            };
            SelectedUnitOption = UnitOptions[1];
            _workbookState = PostprocessingWorkbookStateStore.Load(WorkbookStateKey);
            RestoreWorkbookState();
            _isWorkbookStateLoaded = true;
            LoadCases = new ObservableCollection<BaseReactionOutputCaseViewModel>();
            LoadCombinations = new ObservableCollection<BaseReactionOutputCaseViewModel>();
            LoadOutputCasesCommand = new RelayCommand(LoadOutputCases, () => !IsBusy);
            UseActiveCellCommand = new RelayCommand(UseActiveCell, () => !IsBusy);
            PickAnchorCellCommand = new RelayCommand(() => { IsPickCellMode = true; }, () => !IsBusy);
            RunCommand = new RelayCommand(Run, () => !IsBusy);
            CancelCommand = new RelayCommand(() => RequestClose?.Invoke(this, EventArgs.Empty));

            RefreshAnchorDisplay();
            LoadOutputCases();
        }

        public event EventHandler RequestClose;

        public ObservableCollection<BaseReactionOutputCaseViewModel> LoadCases { get; private set; }

        public ObservableCollection<BaseReactionOutputCaseViewModel> LoadCombinations { get; private set; }

        public ObservableCollection<BaseReactionUnitOption> UnitOptions { get; private set; }

        public ICommand LoadOutputCasesCommand { get; private set; }
        public ICommand UseActiveCellCommand { get; private set; }
        public ICommand PickAnchorCellCommand { get; private set; }
        public ICommand RunCommand { get; private set; }
        public ICommand CancelCommand { get; private set; }

        private BaseReactionUnitOption _selectedUnitOption;
        public BaseReactionUnitOption SelectedUnitOption
        {
            get { return _selectedUnitOption; }
            set
            {
                _selectedUnitOption = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(ForceUnitText));
                OnPropertyChanged(nameof(MomentUnitText));
                SaveWorkbookState();
            }
        }

        public string ForceUnitText
        {
            get { return SelectedUnitOption == null ? string.Empty : $"Force: {SelectedUnitOption.ForceUnit}"; }
        }

        public string MomentUnitText
        {
            get { return SelectedUnitOption == null ? string.Empty : $"Moment: {SelectedUnitOption.MomentUnit}"; }
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

        public bool IsBusy
        {
            get { return _isBusy; }
            private set
            {
                _isBusy = value;
                OnPropertyChanged();
                RaiseCommandStates();
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
                if (!value || _isUseActiveCellMode == value)
                {
                    return;
                }

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
                if (!value || _isPickCellMode == value)
                {
                    return;
                }

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

        public int SelectedLoadCombinationCount
        {
            get { return _selectedLoadCombinationCount; }
            private set
            {
                _selectedLoadCombinationCount = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(LoadCombinationSelectionText));
            }
        }

        public string LoadCaseSelectionText
        {
            get { return $"{SelectedLoadCaseCount}/{LoadCases.Count} selected"; }
        }

        public string LoadCombinationSelectionText
        {
            get { return $"{SelectedLoadCombinationCount}/{LoadCombinations.Count} selected"; }
        }

        public void UpdateSelectionCounts(int selectedLoadCaseCount, int selectedLoadCombinationCount)
        {
            SelectedLoadCaseCount = selectedLoadCaseCount;
            SelectedLoadCombinationCount = selectedLoadCombinationCount;
        }

        public void UpdateSelectedOutputCases(System.Collections.IList selectedLoadCases, System.Collections.IList selectedLoadCombinations)
        {
            _selectedLoadCaseNames = GetSelectedOutputCaseNames(selectedLoadCases);
            _selectedLoadCombinationNames = GetSelectedOutputCaseNames(selectedLoadCombinations);
            UpdateSelectionCounts(_selectedLoadCaseNames.Count, _selectedLoadCombinationNames.Count);
            SaveWorkbookState();
        }

        public void RestoreSavedSelections(System.Collections.IList selectedLoadCases, System.Collections.IList selectedLoadCombinations)
        {
            RestoreSelectedItems(selectedLoadCases, LoadCases, _workbookState.LoadCaseNames);
            RestoreSelectedItems(selectedLoadCombinations, LoadCombinations, _workbookState.LoadCombinationNames);
            UpdateSelectedOutputCases(selectedLoadCases, selectedLoadCombinations);
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
                StatusText = "Loading ETABS load cases and combinations...";
                var result = _csiConnectionService.GetAnalysisOutputCases();
                if (!result.IsSuccess)
                {
                    ShowWarning(result.Message);
                    StatusText = result.Message;
                    return;
                }

                LoadCases.Clear();
                LoadCombinations.Clear();
                if (result.Data != null)
                {
                    foreach (var outputCase in result.Data)
                    {
                        var item = new BaseReactionOutputCaseViewModel(outputCase);
                        if (outputCase.IsLoadCombination)
                        {
                            LoadCombinations.Add(item);
                        }
                        else
                        {
                            LoadCases.Add(item);
                        }
                    }
                }

                int totalCount = LoadCases.Count + LoadCombinations.Count;
                UpdateSelectionCounts(0, 0);
                OnPropertyChanged(nameof(LoadCaseSelectionText));
                OnPropertyChanged(nameof(LoadCombinationSelectionText));
                StatusText = totalCount == 0
                    ? "No ETABS load cases or load combinations were found."
                    : $"Loaded {LoadCases.Count} load case(s) and {LoadCombinations.Count} load combination(s).";
            }
            catch (Exception ex)
            {
                StatusText = "Failed to load ETABS output cases.";
                ShowError($"Failed to load ETABS load cases and combinations: {ex.Message}");
            }
            finally
            {
                IsBusy = false;
            }
        }

        private void UseActiveCell()
        {
            IsUseActiveCellMode = true;
        }

        private void RefreshActiveCellDisplay()
        {
            ExcelRange activeCell = GetActiveExcelCell();
            if (activeCell == null)
            {
                AnchorCellAddress = string.Empty;
                StatusText = "Select an Excel anchor cell for the first data row.";
                return;
            }

            AnchorCellAddress = FormatAddress(activeCell);
            StatusText = $"Anchor cell set to {AnchorCellAddress}.";
        }

        private bool PickAnchorCell()
        {
            try
            {
                var excelApp = ExcelApplicationProvider.GetApplication();
                if (excelApp == null)
                {
                    ShowWarning("Excel application is not available.");
                    return false;
                }

                object result = excelApp.InputBox(
                    AddHeaders
                        ? "Select the top-left anchor cell where Base Reactions headers should start. Data will start one row below."
                        : "Select the top-left anchor cell where the first Base Reactions data row should start. Headers are excluded.",
                    "Get Base Reactions",
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
        }

        public void Run(System.Collections.IList selectedLoadCases, System.Collections.IList selectedLoadCombinations)
        {
            if (!EnsureEtabs())
            {
                return;
            }

            if (!PrepareAnchorCellForWrite())
            {
                return;
            }

            var selectedCases = GetSelectedOutputCases(selectedLoadCases, selectedLoadCombinations);
            if (selectedCases.Count == 0)
            {
                ShowWarning("Select at least one ETABS load case or load combination.");
                return;
            }

            if (!ApplySelectedUnits())
            {
                return;
            }

            try
            {
                IsBusy = true;
                StatusText = "Extracting ETABS Base Reactions...";
                var result = _getBaseReactionsUseCase.Execute(selectedCases);
                if (!result.IsSuccess)
                {
                    StatusText = result.Message;
                    ShowWarning(result.Message);
                    return;
                }

                if (result.Data == null || result.Data.Count == 0)
                {
                    StatusText = "ETABS returned no Base Reactions records.";
                    MessageBox.Show(
                        "ETABS returned no Base Reactions records for the selected cases/combinations. Nothing was written to Excel.",
                        "Get Base Reactions",
                        MessageBoxButton.OK,
                        MessageBoxImage.Information);
                    return;
                }

                object[,] values = CreateOutputValues(result.Data, AddHeaders, SelectedUnitOption);
                OperationResult writeResult = _excelOutputService.WriteValuesToActiveCell(
                    values,
                    $"Successfully wrote {result.Data.Count} Base Reactions record(s) to Excel.",
                    AddHeaders);

                StatusText = writeResult.Message;
                MessageBox.Show(
                    writeResult.Message,
                    "Get Base Reactions",
                    MessageBoxButton.OK,
                    writeResult.IsSuccess ? MessageBoxImage.Information : MessageBoxImage.Warning);
            }
            catch (Exception ex)
            {
                StatusText = "Failed to extract Base Reactions.";
                ShowError($"Failed to extract Base Reactions: {ex.Message}");
            }
            finally
            {
                IsBusy = false;
            }
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

        private bool ApplySelectedUnits()
        {
            if (SelectedUnitOption == null)
            {
                ShowWarning("Select ETABS output units before running.");
                return false;
            }

            try
            {
                OperationResult unitResult = _csiConnectionService.SetPresentUnits(SelectedUnitOption.EtabsUnitsCode);
                if (!unitResult.IsSuccess)
                {
                    ShowWarning(string.IsNullOrWhiteSpace(unitResult.Message)
                        ? "Failed to set ETABS present units."
                        : unitResult.Message);
                    return false;
                }

                return true;
            }
            catch (Exception ex)
            {
                ShowWarning($"Failed to set ETABS present units: {ex.Message}");
                return false;
            }
        }

        private void Run()
        {
            Run(null, null);
        }

        private static List<CSISapModelOutputCaseDTO> GetSelectedOutputCases(
            System.Collections.IList selectedLoadCases,
            System.Collections.IList selectedLoadCombinations)
        {
            var selectedCases = new List<CSISapModelOutputCaseDTO>();
            AddSelectedOutputCases(selectedCases, selectedLoadCases);
            AddSelectedOutputCases(selectedCases, selectedLoadCombinations);
            return selectedCases;
        }

        private static void AddSelectedOutputCases(
            ICollection<CSISapModelOutputCaseDTO> selectedCases,
            System.Collections.IList selectedItems)
        {
            if (selectedCases == null || selectedItems == null)
            {
                return;
            }

            foreach (var selectedItem in selectedItems)
            {
                var item = selectedItem as BaseReactionOutputCaseViewModel;
                if (item != null && item.OutputCase != null)
                {
                    selectedCases.Add(item.OutputCase);
                }
            }
        }

        private static object[,] CreateOutputValues(
            IReadOnlyList<CSISapModelBaseReactionRowDTO> rows,
            bool addHeaders,
            BaseReactionUnitOption unitOption)
        {
            int headerOffset = addHeaders ? 1 : 0;
            var values = new object[rows.Count + headerOffset, 13];

            if (addHeaders)
            {
                string forceUnit = unitOption == null ? string.Empty : unitOption.ForceUnit;
                string momentUnit = unitOption == null ? string.Empty : unitOption.MomentUnit;
                string[] headers = new[]
                {
                    "Output Case",
                    "Case Type",
                    "Step Type",
                    "Step Number",
                    FormatUnitHeader("FX", forceUnit),
                    FormatUnitHeader("FY", forceUnit),
                    FormatUnitHeader("FZ", forceUnit),
                    FormatUnitHeader("MX", momentUnit),
                    FormatUnitHeader("MY", momentUnit),
                    FormatUnitHeader("MZ", momentUnit),
                    "X",
                    "Y",
                    "Z"
                };

                for (int col = 0; col < headers.Length; col++)
                {
                    values[0, col] = headers[col];
                }
            }

            for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++)
            {
                CSISapModelBaseReactionRowDTO row = rows[rowIndex];
                int targetRowIndex = rowIndex + headerOffset;
                values[targetRowIndex, 0] = row.OutputCase;
                values[targetRowIndex, 1] = row.CaseType;
                values[targetRowIndex, 2] = row.StepType;
                values[targetRowIndex, 3] = row.StepNumber;
                values[targetRowIndex, 4] = row.FX;
                values[targetRowIndex, 5] = row.FY;
                values[targetRowIndex, 6] = row.FZ;
                values[targetRowIndex, 7] = row.MX;
                values[targetRowIndex, 8] = row.MY;
                values[targetRowIndex, 9] = row.MZ;
                values[targetRowIndex, 10] = row.X;
                values[targetRowIndex, 11] = row.Y;
                values[targetRowIndex, 12] = row.Z;
            }

            return values;
        }

        private static string FormatUnitHeader(string name, string unit)
        {
            return string.IsNullOrWhiteSpace(unit) ? name : $"{name} ({unit})";
        }

        private bool EnsureEtabs()
        {
            if (!string.Equals(_csiConnectionService.ProductName, "ETABS", StringComparison.OrdinalIgnoreCase))
            {
                ShowWarning("Get Base Reactions is available from the ETABS Toolbox only.");
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

            foreach (BaseReactionUnitOption unitOption in UnitOptions)
            {
                if (string.Equals(unitOption.Label, _workbookState.UnitLabel, StringComparison.OrdinalIgnoreCase))
                {
                    SelectedUnitOption = unitOption;
                    break;
                }
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
                UnitLabel = SelectedUnitOption == null ? string.Empty : SelectedUnitOption.Label,
                AddHeaders = AddHeaders,
                UsePickedAnchor = IsPickCellMode,
                AnchorAddress = IsPickCellMode ? AnchorCellAddress : string.Empty,
                LoadCaseNames = _selectedLoadCaseNames,
                LoadCombinationNames = _selectedLoadCombinationNames
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

        private void RaiseCommandStates()
        {
            RaiseCommandState(LoadOutputCasesCommand);
            RaiseCommandState(UseActiveCellCommand);
            RaiseCommandState(PickAnchorCellCommand);
            RaiseCommandState(RunCommand);
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
                "Get Base Reactions",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
        }

        private static void ShowError(string message)
        {
            MessageBox.Show(
                string.IsNullOrWhiteSpace(message) ? "An unexpected error occurred." : message,
                "Get Base Reactions",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
        }
    }

    public class BaseReactionUnitOption
    {
        public BaseReactionUnitOption(string label, int etabsUnitsCode, string forceUnit, string momentUnit, string lengthUnit)
        {
            Label = label;
            EtabsUnitsCode = etabsUnitsCode;
            ForceUnit = forceUnit;
            MomentUnit = momentUnit;
            LengthUnit = lengthUnit;
        }

        public string Label { get; private set; }

        public int EtabsUnitsCode { get; private set; }

        public string ForceUnit { get; private set; }

        public string MomentUnit { get; private set; }

        public string LengthUnit { get; private set; }
    }
}
