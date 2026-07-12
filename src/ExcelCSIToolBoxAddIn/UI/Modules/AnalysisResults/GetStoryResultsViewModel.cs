using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Input;
using ExcelCSIToolBox.Application.Services;
using ExcelCSIToolBox.Application.Composition;
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
    public class GetStoryResultsViewModel : ViewModelBase
    {
        private readonly CsiToolboxUseCaseBundle _useCases;
        private readonly ICSISapModelConnectionService _csiConnectionService;
        private readonly IExcelOutputService _excelOutputService;
        private readonly StoryPostprocessingResultKind _kind;
        private BaseReactionUnitOption _selectedUnitOption;
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

        public GetStoryResultsViewModel(
            StoryPostprocessingResultKind kind,
            CsiToolboxUseCaseBundle useCases,
            ICSISapModelConnectionService csiConnectionService,
            IExcelOutputService excelOutputService,
            BaseReactionUnitOption exportUnitOption = null)
        {
            _kind = kind;
            _useCases = useCases ?? throw new ArgumentNullException(nameof(useCases));
            _csiConnectionService = csiConnectionService ?? throw new ArgumentNullException(nameof(csiConnectionService));
            _excelOutputService = excelOutputService ?? throw new ArgumentNullException(nameof(excelOutputService));

            UnitOptions = new ObservableCollection<BaseReactionUnitOption>
            {
                new BaseReactionUnitOption("N-mm", 9, "N", "N-mm", "mm"),
                new BaseReactionUnitOption("kN-m", 6, "kN", "kN-m", "m"),
                new BaseReactionUnitOption("kip-ft", 4, "kip", "kip-ft", "ft"),
                new BaseReactionUnitOption("lb-in", 1, "lb", "lb-in", "in")
            };
            SelectedUnitOption = exportUnitOption ?? UnitOptions[1];
            _workbookState = PostprocessingWorkbookStateStore.Load(GetWorkbookStateKey());
            RestoreWorkbookState();
            _isWorkbookStateLoaded = true;
            LoadCases = new ObservableCollection<BaseReactionOutputCaseViewModel>();
            LoadCombinations = new ObservableCollection<BaseReactionOutputCaseViewModel>();

            LoadOutputCasesCommand = new RelayCommand(LoadOutputCases, () => !IsBusy);
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
        public ICommand PickAnchorCellCommand { get; private set; }
        public ICommand RunCommand { get; private set; }
        public ICommand CancelCommand { get; private set; }

        public string WindowTitle
        {
            get
            {
                switch (_kind)
                {
                    case StoryPostprocessingResultKind.StoryDrifts:
                        return "Story Drifts";
                    case StoryPostprocessingResultKind.StoryMaxOverAverageDisplacements:
                        return "Story Max Over Avg Displacements";
                    case StoryPostprocessingResultKind.StoryMaxOverAverageDrifts:
                        return "Story Max Over Avg Drifts";
                    default:
                        return "Story Forces";
                }
            }
        }

        public string InstructionText
        {
            get { return $"Select ETABS load cases / combinations to extract {WindowTitle}, then select the Excel anchor cell where output should start."; }
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

        public BaseReactionUnitOption SelectedUnitOption
        {
            get { return _selectedUnitOption; }
            set
            {
                _selectedUnitOption = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(ForceUnitText));
                OnPropertyChanged(nameof(MomentUnitText));
                OnPropertyChanged(nameof(LengthUnitText));
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

        public string LengthUnitText
        {
            get { return SelectedUnitOption == null ? string.Empty : $"Length: {SelectedUnitOption.LengthUnit}"; }
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
            get { return IsPickCellMode ? "Picked cell is fixed until changed." : "Uses the current Excel active cell when Run is clicked."; }
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
                    StatusText = result.Message;
                    ShowWarning(result.Message);
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

                UpdateSelectionCounts(0, 0);
                OnPropertyChanged(nameof(LoadCaseSelectionText));
                OnPropertyChanged(nameof(LoadCombinationSelectionText));
                StatusText = $"Loaded {LoadCases.Count} load case(s) and {LoadCombinations.Count} load combination(s).";
            }
            catch (Exception ex)
            {
                StatusText = $"Failed to load ETABS output cases.";
                ShowError($"Failed to load ETABS output cases: {ex.Message}");
            }
            finally
            {
                IsBusy = false;
            }
        }

        public void Run(System.Collections.IList selectedLoadCases, System.Collections.IList selectedLoadCombinations)
        {
            if (!EnsureEtabs() || !PrepareAnchorCellForWrite())
            {
                return;
            }

            CsiPresentUnitScope unitScope = null;
            try
            {
                RaiseRequestHide();

                var selectedCases = GetSelectedOutputCases(selectedLoadCases, selectedLoadCombinations);
                if (selectedCases.Count == 0)
                {
                    ShowWarning("Select at least one ETABS load case or load combination.");
                    return;
                }

                OperationResult<CsiPresentUnitScope> unitScopeResult = ApplySelectedUnitScope();
                if (!unitScopeResult.IsSuccess)
                {
                    return;
                }

                unitScope = unitScopeResult.Data;
                if (_kind == StoryPostprocessingResultKind.StoryForces)
                {
                    RunStoryForces(selectedCases);
                }
                else
                {
                    RunStoryTable(selectedCases);
                }
            }
            finally
            {
                RestoreSelectedUnitScope(unitScope);
                RaiseRequestShow();
            }
        }

        private void RunStoryForces(IReadOnlyList<CSISapModelOutputCaseDTO> selectedCases)
        {
            try
            {
                IsBusy = true;
                StatusText = "Extracting ETABS Story Forces...";
                var result = _useCases.GetStoryForces.Execute(selectedCases);
                if (!result.IsSuccess)
                {
                    StatusText = result.Message;
                    ShowWarning(result.Message);
                    return;
                }

                if (result.Data == null || result.Data.Count == 0)
                {
                    ShowNoRecordsMessage();
                    return;
                }

                object[,] values = CreateStoryForceOutputValues(result.Data, AddHeaders, SelectedUnitOption);
                OperationResult writeResult = _excelOutputService.WriteValuesToActiveCell(
                    values,
                    $"Successfully wrote {result.Data.Count} Story Forces record(s) to Excel.",
                    AddHeaders);
                ShowWriteResult(writeResult);
            }
            catch (Exception ex)
            {
                StatusText = "Failed to extract Story Forces.";
                ShowError($"Failed to extract Story Forces: {ex.Message}");
            }
            finally
            {
                IsBusy = false;
            }
        }

        private void RunStoryTable(IReadOnlyList<CSISapModelOutputCaseDTO> selectedCases)
        {
            try
            {
                IsBusy = true;
                StatusText = $"Extracting ETABS {WindowTitle}...";
                OperationResult<CSISapModelDisplayTableDTO> result;
                switch (_kind)
                {
                    case StoryPostprocessingResultKind.StoryDrifts:
                        result = _useCases.GetStoryDrifts.Execute(selectedCases);
                        break;
                    case StoryPostprocessingResultKind.StoryMaxOverAverageDisplacements:
                        result = _useCases.GetStoryMaxOverAverageDisplacements.Execute(selectedCases);
                        break;
                    case StoryPostprocessingResultKind.StoryMaxOverAverageDrifts:
                        result = _useCases.GetStoryMaxOverAverageDrifts.Execute(selectedCases);
                        break;
                    default:
                        ShowWarning("Select a supported ETABS story result table.");
                        return;
                }

                if (!result.IsSuccess)
                {
                    StatusText = result.Message;
                    ShowWarning(result.Message);
                    return;
                }

                if (result.Data == null || result.Data.Rows == null || result.Data.Rows.Count == 0)
                {
                    ShowNoRecordsMessage();
                    return;
                }

                object[,] values = CreateStoryTableOutputValues(result.Data, AddHeaders, SelectedUnitOption);
                OperationResult writeResult = _excelOutputService.WriteValuesToActiveCell(
                    values,
                    $"Successfully wrote {result.Data.Rows.Count} {WindowTitle} record(s) to Excel.",
                    AddHeaders);
                ShowWriteResult(writeResult);
            }
            catch (Exception ex)
            {
                StatusText = $"Failed to extract {WindowTitle}.";
                ShowError($"Failed to extract {WindowTitle}: {ex.Message}");
            }
            finally
            {
                IsBusy = false;
            }
        }

        private void ShowNoRecordsMessage()
        {
            StatusText = $"ETABS returned no {WindowTitle} records.";
            MessageBox.Show(
                $"ETABS returned no {WindowTitle} records for the selected cases/combinations. Nothing was written to Excel.",
                WindowTitle,
                MessageBoxButton.OK,
                MessageBoxImage.Information);
        }

        private void ShowWriteResult(OperationResult writeResult)
        {
            StatusText = writeResult.Message;
            MessageBox.Show(
                writeResult.Message,
                WindowTitle,
                MessageBoxButton.OK,
                writeResult.IsSuccess ? MessageBoxImage.Information : MessageBoxImage.Warning);
        }

        private void Run()
        {
            Run(null, null);
        }

        private OperationResult<CsiPresentUnitScope> ApplySelectedUnitScope()
        {
            if (SelectedUnitOption == null)
            {
                ShowWarning("Select ETABS output units before running.");
                return OperationResult<CsiPresentUnitScope>.Failure("Select ETABS output units before running.");
            }

            try
            {
                OperationResult<CsiPresentUnitScope> unitScopeResult = CsiPresentUnitScope.Apply(
                    _csiConnectionService,
                    SelectedUnitOption.ToPresentUnitSystemDto());
                if (!unitScopeResult.IsSuccess)
                {
                    ShowWarning(string.IsNullOrWhiteSpace(unitScopeResult.Message)
                        ? "Failed to set ETABS present units."
                        : unitScopeResult.Message);
                    return unitScopeResult;
                }

                return unitScopeResult;
            }
            catch (Exception ex)
            {
                ShowWarning($"Failed to set ETABS present units: {ex.Message}");
                return OperationResult<CsiPresentUnitScope>.Failure("Failed to set ETABS present units: " + ex.Message);
            }
        }

        private void RestoreSelectedUnitScope(CsiPresentUnitScope unitScope)
        {
            if (unitScope == null)
            {
                return;
            }

            unitScope.Dispose();
            if (unitScope.RestoreResult != null && !unitScope.RestoreResult.IsSuccess)
            {
                AnalysisExportDiagnostics.Log(
                    "Failed to restore ETABS present units after " + WindowTitle + ": " + unitScope.RestoreResult.Message);
            }
        }

        private static List<CSISapModelOutputCaseDTO> GetSelectedOutputCases(System.Collections.IList selectedLoadCases, System.Collections.IList selectedLoadCombinations)
        {
            var selectedCases = new List<CSISapModelOutputCaseDTO>();
            AddSelectedOutputCases(selectedCases, selectedLoadCases);
            AddSelectedOutputCases(selectedCases, selectedLoadCombinations);
            return selectedCases;
        }

        private static void AddSelectedOutputCases(ICollection<CSISapModelOutputCaseDTO> selectedCases, System.Collections.IList selectedItems)
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

        private static object[,] CreateStoryForceOutputValues(IReadOnlyList<CSISapModelStoryForceRowDTO> rows, bool addHeaders, BaseReactionUnitOption unitOption)
        {
            int headerOffset = addHeaders ? 1 : 0;
            var values = new object[rows.Count + headerOffset, 12];
            if (addHeaders)
            {
                string forceUnit = unitOption == null ? string.Empty : unitOption.ForceUnit;
                string momentUnit = unitOption == null ? string.Empty : unitOption.MomentUnit;
                string[] headers =
                {
                    "Story",
                    "Output Case",
                    "Case Type",
                    "Step Type",
                    "Step Number",
                    "Location",
                    FormatUnitHeader("P", forceUnit),
                    FormatUnitHeader("VX", forceUnit),
                    FormatUnitHeader("VY", forceUnit),
                    FormatUnitHeader("T", momentUnit),
                    FormatUnitHeader("MX", momentUnit),
                    FormatUnitHeader("MY", momentUnit)
                };
                for (int col = 0; col < headers.Length; col++) values[0, col] = headers[col];
            }

            for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++)
            {
                var row = rows[rowIndex];
                int targetRowIndex = rowIndex + headerOffset;
                values[targetRowIndex, 0] = row.Story;
                values[targetRowIndex, 1] = row.OutputCase;
                values[targetRowIndex, 2] = row.CaseType;
                values[targetRowIndex, 3] = row.StepType;
                values[targetRowIndex, 4] = row.StepNumber;
                values[targetRowIndex, 5] = row.Location;
                values[targetRowIndex, 6] = row.P;
                values[targetRowIndex, 7] = row.VX;
                values[targetRowIndex, 8] = row.VY;
                values[targetRowIndex, 9] = row.T;
                values[targetRowIndex, 10] = row.MX;
                values[targetRowIndex, 11] = row.MY;
            }

            return values;
        }

        private static object[,] CreateStoryDisplacementOutputValues(IReadOnlyList<CSISapModelStoryDisplacementRowDTO> rows, bool addHeaders, BaseReactionUnitOption unitOption)
        {
            int headerOffset = addHeaders ? 1 : 0;
            var values = new object[rows.Count + headerOffset, 11];
            if (addHeaders)
            {
                string lengthUnit = unitOption == null ? string.Empty : unitOption.LengthUnit;
                string[] headers =
                {
                    "Story",
                    "Output Case",
                    "Case Type",
                    "Step Type",
                    "Step Number",
                    FormatUnitHeader("UX", lengthUnit),
                    FormatUnitHeader("UY", lengthUnit),
                    FormatUnitHeader("UZ", lengthUnit),
                    "RX (rad)",
                    "RY (rad)",
                    "RZ (rad)"
                };
                for (int col = 0; col < headers.Length; col++) values[0, col] = headers[col];
            }

            for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++)
            {
                var row = rows[rowIndex];
                int targetRowIndex = rowIndex + headerOffset;
                values[targetRowIndex, 0] = row.Story;
                values[targetRowIndex, 1] = row.OutputCase;
                values[targetRowIndex, 2] = row.CaseType;
                values[targetRowIndex, 3] = row.StepType;
                values[targetRowIndex, 4] = row.StepNumber;
                values[targetRowIndex, 5] = row.UX;
                values[targetRowIndex, 6] = row.UY;
                values[targetRowIndex, 7] = row.UZ;
                values[targetRowIndex, 8] = row.RX;
                values[targetRowIndex, 9] = row.RY;
                values[targetRowIndex, 10] = row.RZ;
            }

            return values;
        }

        private static object[,] CreateStoryTableOutputValues(CSISapModelDisplayTableDTO table, bool addHeaders, BaseReactionUnitOption unitOption)
        {
            int fieldCount = table.FieldKeys == null ? 0 : table.FieldKeys.Count;
            int rowCount = table.Rows == null ? 0 : table.Rows.Count;
            int headerOffset = addHeaders ? 1 : 0;
            var values = new object[rowCount + headerOffset, fieldCount];

            if (addHeaders)
            {
                for (int fieldIndex = 0; fieldIndex < fieldCount; fieldIndex++)
                {
                    values[0, fieldIndex] = FormatStoryTableHeader(table.FieldKeys[fieldIndex], unitOption);
                }
            }

            for (int rowIndex = 0; rowIndex < rowCount; rowIndex++)
            {
                object[] row = table.Rows[rowIndex];
                for (int fieldIndex = 0; fieldIndex < fieldCount; fieldIndex++)
                {
                    values[rowIndex + headerOffset, fieldIndex] = row != null && fieldIndex < row.Length
                        ? row[fieldIndex]
                        : string.Empty;
                }
            }

            return values;
        }

        private static string FormatStoryTableHeader(string fieldKey, BaseReactionUnitOption unitOption)
        {
            if (string.IsNullOrWhiteSpace(fieldKey) || unitOption == null)
            {
                return fieldKey ?? string.Empty;
            }

            return fieldKey.IndexOf("displ", StringComparison.OrdinalIgnoreCase) >= 0
                ? FormatUnitHeader(fieldKey, unitOption.LengthUnit)
                : fieldKey;
        }

        private static string FormatUnitHeader(string name, string unit)
        {
            return string.IsNullOrWhiteSpace(unit) ? name : $"{name} ({unit})";
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
                        ? $"Select the top-left anchor cell where {WindowTitle} headers should start. Data will start one row below."
                        : $"Select the top-left anchor cell where the first {WindowTitle} data row should start. Headers are excluded.",
                    WindowTitle,
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
                ShowWarning($"{WindowTitle} is available from the ETABS Toolbox only.");
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

        private string GetWorkbookStateKey()
        {
            return "Story." + _kind;
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

            PostprocessingWorkbookStateStore.Save(GetWorkbookStateKey(), new PostprocessingWorkbookState
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

        private static void RaiseCommandState(ICommand command)
        {
            var relayCommand = command as IRelayCommand;
            if (relayCommand != null)
            {
                relayCommand.RaiseCanExecuteChanged();
            }
        }

        private void ShowWarning(string message)
        {
            MessageBox.Show(
                string.IsNullOrWhiteSpace(message) ? "The operation could not be completed." : message,
                WindowTitle,
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
        }

        private void ShowError(string message)
        {
            MessageBox.Show(
                string.IsNullOrWhiteSpace(message) ? "An unexpected error occurred." : message,
                WindowTitle,
                MessageBoxButton.OK,
                MessageBoxImage.Error);
        }
    }
}
