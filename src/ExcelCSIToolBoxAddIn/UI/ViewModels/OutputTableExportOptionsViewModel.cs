using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Input;
using ExcelCSIToolBox.Application.Services;
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
    public class OutputTableExportOptionsViewModel : ViewModelBase
    {
        private readonly ICSISapModelConnectionService _csiConnectionService;
        private readonly IExcelOutputService _excelOutputService;
        private readonly GetBaseReactionsUseCase _getBaseReactionsUseCase;
        private readonly OutputTableExportConfig _config;
        private readonly OutputTablePopupProfile _profile;
        private readonly string _displayTableName;
        private readonly string _workbookStateKey;
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
        private bool _hasInitializedForDialog;
        private string _etabsModelName = "ETABS Model: Not attached";

        public OutputTableExportOptionsViewModel(
            CsiToolboxUseCaseBundle useCases,
            ICSISapModelConnectionService csiConnectionService,
            IExcelOutputService excelOutputService)
            : this(
                useCases,
                csiConnectionService,
                excelOutputService,
                new OutputTableExportConfig
                {
                    TableDisplayName = "Base Reactions",
                    Breadcrumb = "ETABS Toolbox / ANALYSIS RESULTS / Base Reactions"
                })
        {
        }

        public OutputTableExportOptionsViewModel(
            CsiToolboxUseCaseBundle useCases,
            ICSISapModelConnectionService csiConnectionService,
            IExcelOutputService excelOutputService,
            string displayTableName)
            : this(
                useCases,
                csiConnectionService,
                excelOutputService,
                OutputTableExportConfig.ForTable(
                    displayTableName,
                    "ETABS Toolbox / ANALYSIS RESULTS / " + (string.IsNullOrWhiteSpace(displayTableName) ? "Base Reactions" : displayTableName)))
        {
        }

        public OutputTableExportOptionsViewModel(
            CsiToolboxUseCaseBundle useCases,
            ICSISapModelConnectionService csiConnectionService,
            IExcelOutputService excelOutputService,
            OutputTableExportConfig config)
        {
            if (useCases == null) throw new ArgumentNullException(nameof(useCases));
            _csiConnectionService = csiConnectionService ?? throw new ArgumentNullException(nameof(csiConnectionService));
            _excelOutputService = excelOutputService ?? throw new ArgumentNullException(nameof(excelOutputService));
            _getBaseReactionsUseCase = useCases.GetBaseReactions ?? throw new ArgumentNullException(nameof(useCases.GetBaseReactions));
            _config = (config ?? new OutputTableExportConfig()).Normalize();
            _profile = OutputTablePopupProfileProvider.GetProfile(_config.PopupProfileKey);
            _displayTableName = _config.TableDisplayName;
            _workbookStateKey = "OutputTableExport." + CreateStateKey(_profile.WorksheetNamePrefix + "." + _config.TableDisplayName);
            _addHeaders = _config.DefaultAddHeaders;

            UnitOptions = new ObservableCollection<BaseReactionUnitOption>
            {
                new BaseReactionUnitOption("N-mm", 9, "N", "N-mm", "mm"),
                new BaseReactionUnitOption("kN-m", 6, "kN", "kN-m", "m"),
                new BaseReactionUnitOption("kip-ft", 4, "kip", "kip-ft", "ft"),
                new BaseReactionUnitOption("lb-in", 1, "lb", "lb-in", "in")
            };
            SelectedUnitOption = FindUnitOption(_config.ExportUnitOption) ?? UnitOptions[1];
            _workbookState = PostprocessingWorkbookStateStore.Load(_workbookStateKey);
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
        }

        public event EventHandler RequestClose;

        public string WindowTitle
        {
            get { return "Export Analysis Report"; }
        }

        public string Breadcrumb
        {
            get { return _config.Breadcrumb; }
        }

        public string Description
        {
            get { return _config.Description; }
        }

        public string ReportType
        {
            get { return _displayTableName; }
        }

        public string ReportTypeText
        {
            get { return "Report: " + _displayTableName; }
        }

        public Visibility DescriptionVisibility
        {
            get { return string.IsNullOrWhiteSpace(Description) ? Visibility.Collapsed : Visibility.Visible; }
        }

        public Visibility CaseComboSelectorVisibility
        {
            get { return _profile.ShowCaseComboSelector ? Visibility.Visible : Visibility.Collapsed; }
        }

        public GridLength CaseComboSelectorRowHeight
        {
            get { return _profile.ShowCaseComboSelector ? new GridLength(1, GridUnitType.Star) : GridLength.Auto; }
        }

        public Visibility UnitSelectorVisibility
        {
            get { return _profile.ShowUnitSelector ? Visibility.Visible : Visibility.Collapsed; }
        }

        public bool AllowMultipleCases
        {
            get { return _profile.AllowMultipleCases; }
        }

        public string CaseSelectorTitle
        {
            get { return string.IsNullOrWhiteSpace(_profile.CaseSelectorTitle) ? "Load Case" : _profile.CaseSelectorTitle; }
        }

        public Visibility LoadCombinationSelectorVisibility
        {
            get { return _profile.ShowComboSelector ? Visibility.Visible : Visibility.Collapsed; }
        }

        public GridLength LoadCombinationColumnWidth
        {
            get { return _profile.ShowComboSelector ? new GridLength(1, GridUnitType.Star) : new GridLength(0); }
        }

        public GridLength LoadCombinationSpacerWidth
        {
            get { return _profile.ShowComboSelector ? new GridLength(10) : new GridLength(0); }
        }

        public string OutputCaseSelectorHelpText
        {
            get
            {
                if (!_profile.ShowComboSelector)
                {
                    return _profile.AllowMultipleCases
                        ? "Choose one or more " + CaseSelectorTitle.ToLowerInvariant() + " rows."
                        : "Choose one " + CaseSelectorTitle.ToLowerInvariant() + ".";
                }

                return _profile.AllowMultipleCases
                    ? "Choose one or more rows from either list."
                    : "Choose one " + CaseSelectorTitle.ToLowerInvariant() + ".";
            }
        }

        public string InstructionText
        {
            get
            {
                if (!_profile.ShowCaseComboSelector)
                {
                    return "Select the Excel anchor cell where " + _displayTableName + " output should start.";
                }

                return "Select ETABS " + CaseSelectorTitle.ToLowerInvariant() + " to extract " + _displayTableName + ", then select the Excel anchor cell where output should start.";
            }
        }

        public ObservableCollection<BaseReactionOutputCaseViewModel> LoadCases { get; private set; }

        public ObservableCollection<BaseReactionOutputCaseViewModel> LoadCombinations { get; private set; }

        public ObservableCollection<BaseReactionUnitOption> UnitOptions { get; private set; }

        public IReadOnlyList<CSISapModelOutputCaseDTO> SelectedCasesOrCombos { get; private set; }

        public BaseReactionUnitOption SelectedUnit
        {
            get { return SelectedUnitOption; }
        }

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
                OnPropertyChanged(nameof(SelectedUnit));
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

        public void RestoreSavedSelections(
            Action<IEnumerable<BaseReactionOutputCaseViewModel>> selectLoadCases,
            Action<IEnumerable<BaseReactionOutputCaseViewModel>> selectLoadCombinations)
        {
            IReadOnlyList<string> loadCaseNames = _workbookState.LoadCaseNames;
            IReadOnlyList<string> loadCombinationNames = _workbookState.LoadCombinationNames;
            if ((loadCaseNames == null || loadCaseNames.Count == 0) &&
                (loadCombinationNames == null || loadCombinationNames.Count == 0) &&
                !string.IsNullOrWhiteSpace(_config.DefaultSelectedCaseOrCombo))
            {
                loadCaseNames = new[] { _config.DefaultSelectedCaseOrCombo };
                loadCombinationNames = new[] { _config.DefaultSelectedCaseOrCombo };
            }

            var casesToSelect = FilterItems(LoadCases, loadCaseNames);
            var combinationsToSelect = FilterItems(LoadCombinations, loadCombinationNames);

            selectLoadCases(casesToSelect);
            selectLoadCombinations(combinationsToSelect);
        }

        private static List<BaseReactionOutputCaseViewModel> FilterItems(
            IEnumerable<BaseReactionOutputCaseViewModel> availableItems,
            IReadOnlyList<string> selectedNames)
        {
            var result = new List<BaseReactionOutputCaseViewModel>();
            if (availableItems == null || selectedNames == null)
            {
                return result;
            }

            var nameSet = new HashSet<string>(selectedNames, StringComparer.OrdinalIgnoreCase);
            foreach (BaseReactionOutputCaseViewModel item in availableItems)
            {
                if (item != null && nameSet.Contains(item.Name))
                {
                    result.Add(item);
                }
            }

            return result;
        }

        public void RefreshAnchorDisplay()
        {
            if (IsUseActiveCellMode)
            {
                RefreshActiveCellDisplay();
            }
        }

        public void InitializeForDialog()
        {
            if (_hasInitializedForDialog)
            {
                return;
            }

            _hasInitializedForDialog = true;
            AnalysisExportDiagnostics.Log("Export options popup rendered: " + _displayTableName);

            if (_profile.ShowUnitSelector)
            {
                SelectCurrentEtabsUnitIfRequested();
            }

            if (_profile.ShowCaseComboSelector)
            {
                LoadOutputCases();
            }
        }

        private void LoadOutputCases()
        {
            AnalysisExportDiagnostics.Log("Loading ETABS load cases and combinations for " + _displayTableName);
            if (!EnsureEtabs())
            {
                AnalysisExportDiagnostics.Log("Unable to load output cases because ETABS is not attached: " + _displayTableName);
                return;
            }

            try
            {
                IsBusy = true;
                StatusText = "Loading ETABS load cases and combinations...";
                OperationResult<IReadOnlyList<CSISapModelOutputCaseDTO>> result = LoadOutputCaseSource();
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
                            if (_profile.CaseSelectionMode != OutputCaseSelectionMode.CaseOnly &&
                                _profile.CaseSelectionMode != OutputCaseSelectionMode.ModalCaseOnly &&
                                _profile.CaseSelectionMode != OutputCaseSelectionMode.SeismicWindOrResponseSpectrumCasesOnly &&
                                _profile.CaseSelectionMode != OutputCaseSelectionMode.None)
                            {
                                LoadCombinations.Add(item);
                            }
                        }
                        else
                        {
                            if (_profile.CaseSelectionMode != OutputCaseSelectionMode.ComboOnly &&
                                _profile.CaseSelectionMode != OutputCaseSelectionMode.None)
                            {
                                if (_profile.CaseSelectionMode == OutputCaseSelectionMode.SeismicWindOrResponseSpectrumCasesOnly)
                                {
                                    if (outputCase.IsSeismicWindOrResponseSpectrum)
                                    {
                                        LoadCases.Add(item);
                                    }
                                }
                                else
                                {
                                    LoadCases.Add(item);
                                }
                            }
                        }
                    }
                }

                int totalCount = LoadCases.Count + LoadCombinations.Count;
                UpdateSelectionCounts(0, 0);
                OnPropertyChanged(nameof(LoadCaseSelectionText));
                OnPropertyChanged(nameof(LoadCombinationSelectionText));
                StatusText = totalCount == 0
                    ? "No ETABS " + CaseSelectorTitle.ToLowerInvariant() + " records were found."
                    : _profile.ShowComboSelector
                        ? $"Loaded {LoadCases.Count} load case(s) and {LoadCombinations.Count} load combination(s)."
                        : $"Loaded {LoadCases.Count} {CaseSelectorTitle.ToLowerInvariant()}(s).";
                AnalysisExportDiagnostics.Log(
                    $"Loaded {LoadCases.Count} load cases and {LoadCombinations.Count} load combinations for {_displayTableName}");
            }
            catch (Exception ex)
            {
                StatusText = "Failed to load ETABS output cases.";
                AnalysisExportDiagnostics.Log("Failed to load ETABS output cases for " + _displayTableName + ": " + ex.Message);
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

        private OperationResult<IReadOnlyList<CSISapModelOutputCaseDTO>> LoadOutputCaseSource()
        {
            if (_profile.CaseSelectionMode == OutputCaseSelectionMode.ModalCaseOnly)
            {
                return _csiConnectionService.GetModalOutputCases();
            }

            if (_profile.CaseSelectionMode == OutputCaseSelectionMode.ResponseSpectrumCaseOnly)
            {
                return _csiConnectionService.GetResponseSpectrumOutputCases();
            }

            return _csiConnectionService.GetAnalysisOutputCases();
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
                RaiseRequestHide();
                var excelApp = ExcelApplicationProvider.GetApplication();
                if (excelApp == null)
                {
                    ShowWarning("Excel application is not available.");
                    return false;
                }

                object result = excelApp.InputBox(
                    AddHeaders
                        ? "Select the top-left anchor cell where " + _displayTableName + " headers should start. Data will start one row below."
                        : "Select the top-left anchor cell where the first " + _displayTableName + " data row should start. Headers are excluded.",
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

        public void Run(System.Collections.IList selectedLoadCases, System.Collections.IList selectedLoadCombinations)
        {
            AnalysisExportDiagnostics.Log("Export confirmation clicked: " + _displayTableName);
            AnalysisExportDiagnostics.Log("Popup result: Confirmed for " + _displayTableName);
            if (!EnsureEtabs())
            {
                AnalysisExportDiagnostics.Log("Export confirmation blocked because ETABS is not attached: " + _displayTableName);
                return;
            }

            if (!PrepareAnchorCellForWrite())
            {
                return;
            }

            CsiPresentUnitScope unitScope = null;
            try
            {
                RaiseRequestHide();

                var selectedCases = _profile.ShowCaseComboSelector
                    ? GetSelectedOutputCases(selectedLoadCases, selectedLoadCombinations)
                    : new List<CSISapModelOutputCaseDTO>();
                if (_profile.ShowCaseComboSelector && selectedCases.Count == 0)
                {
                    AnalysisExportDiagnostics.Log("Export confirmation blocked because no output case was selected: " + _displayTableName);
                    ShowWarning("Select at least one " + CaseSelectorTitle.ToLowerInvariant() + ".");
                    return;
                }

                if (_profile.ShowCaseComboSelector && !_profile.AllowMultipleCases && selectedCases.Count > 1)
                {
                    AnalysisExportDiagnostics.Log("Export confirmation blocked because too many output cases were selected: " + _displayTableName);
                    ShowWarning("Select only one " + CaseSelectorTitle.ToLowerInvariant() + ".");
                    return;
                }

                if (_profile.ShowUnitSelector)
                {
                    OperationResult<CsiPresentUnitScope> unitScopeResult = ApplySelectedUnitScope();
                    if (!unitScopeResult.IsSuccess)
                    {
                        AnalysisExportDiagnostics.Log("Export confirmation blocked because selected units could not be applied: " + _displayTableName);
                        return;
                    }

                    unitScope = unitScopeResult.Data;
                }

                SelectedCasesOrCombos = selectedCases;
                OnPropertyChanged(nameof(SelectedCasesOrCombos));

                try
                {
                    IsBusy = true;
                    StatusText = "Extracting ETABS " + _displayTableName + "...";
                    AnalysisExportDiagnostics.Log(
                        "Starting report export: " + _displayTableName +
                        "; selected cases/combinations=" + selectedCases.Count +
                        "; selected unit=" + (SelectedUnitOption == null ? string.Empty : SelectedUnitOption.Label));
                    if (_config.StaticExportValuesFactory != null)
                    {
                        RunStaticTableExport();
                    }
                    else if (IsBaseReactionsTable())
                    {
                        RunBaseReactionsExport(selectedCases);
                    }
                    else
                    {
                        RunDisplayTableExport(selectedCases);
                    }
                }
                catch (Exception ex)
                {
                    StatusText = "Failed to extract " + _displayTableName + ".";
                    AnalysisExportDiagnostics.Log("Failed to extract " + _displayTableName + ": " + ex.Message);
                    ShowError("Failed to extract " + _displayTableName + ": " + ex.Message);
                }
                finally
                {
                    IsBusy = false;
                    RestoreSelectedUnitScope(unitScope);
                }
            }
            finally
            {
                RaiseRequestShow();
            }
        }

        private void RunStaticTableExport()
        {
            object[,] values = _config.StaticExportValuesFactory(AddHeaders);
            int rowCount = values == null ? 0 : values.GetLength(0);
            int recordCount = _config.StaticRecordCount > 0
                ? _config.StaticRecordCount
                : Math.Max(0, rowCount - (AddHeaders ? 1 : 0));

            if (values == null || values.Length == 0 || rowCount == 0)
            {
                string emptyMessage = string.IsNullOrWhiteSpace(_config.EmptyDataMessage)
                    ? _profile.EmptyDataMessage
                    : _config.EmptyDataMessage;
                StatusText = emptyMessage;
                MessageBox.Show(
                    emptyMessage + " Nothing was written to Excel.",
                    WindowTitle,
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
                AnalysisExportDiagnostics.Log("Static export returned no records for " + _displayTableName);
                return;
            }

            string successMessage = string.IsNullOrWhiteSpace(_config.StaticSuccessMessage)
                ? "Successfully wrote " + recordCount + " " + _displayTableName + " record(s) to Excel."
                : _config.StaticSuccessMessage;
            OperationResult writeResult = _excelOutputService.WriteValuesToActiveCell(
                values,
                successMessage,
                AddHeaders);

            StatusText = writeResult.Message;
            AnalysisExportDiagnostics.Log("Static export completed for " + _displayTableName + ": " + writeResult.Message);
            MessageBox.Show(
                writeResult.Message,
                WindowTitle,
                MessageBoxButton.OK,
                writeResult.IsSuccess ? MessageBoxImage.Information : MessageBoxImage.Warning);
        }

        private void RunBaseReactionsExport(IReadOnlyList<CSISapModelOutputCaseDTO> selectedCases)
        {
            OperationResult<IReadOnlyList<CSISapModelBaseReactionRowDTO>> result = _getBaseReactionsUseCase.Execute(selectedCases);
            if (!result.IsSuccess)
            {
                StatusText = result.Message;
                ShowWarning(result.Message);
                AnalysisExportDiagnostics.Log("Base reactions export failed for " + _displayTableName + ": " + result.Message);
                return;
            }

            if (result.Data == null || result.Data.Count == 0)
            {
                StatusText = "ETABS returned no " + _displayTableName + " records.";
                MessageBox.Show(
                    _profile.EmptyDataMessage + " Nothing was written to Excel.",
                    WindowTitle,
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
                AnalysisExportDiagnostics.Log("Base reactions export returned no records for " + _displayTableName);
                return;
            }

            object[,] values = CreateOutputValues(result.Data, AddHeaders, SelectedUnitOption);
            OperationResult writeResult = _excelOutputService.WriteValuesToActiveCell(
                values,
                "Successfully wrote " + result.Data.Count + " " + _displayTableName + " record(s) to Excel.",
                AddHeaders);

            StatusText = writeResult.Message;
            AnalysisExportDiagnostics.Log("Report export completed for " + _displayTableName + ": " + writeResult.Message);
            MessageBox.Show(
                writeResult.Message,
                WindowTitle,
                MessageBoxButton.OK,
                writeResult.IsSuccess ? MessageBoxImage.Information : MessageBoxImage.Warning);
        }

        private void RunDisplayTableExport(IReadOnlyList<CSISapModelOutputCaseDTO> selectedCases)
        {
            OperationResult<CSISapModelDisplayTableDTO> result = _csiConnectionService.GetDisplayTable(_displayTableName, selectedCases);
            if (!result.IsSuccess)
            {
                StatusText = result.Message;
                ShowWarning(result.Message);
                AnalysisExportDiagnostics.Log("Display table export failed for " + _displayTableName + ": " + result.Message);
                return;
            }

            object[,] values = CreateDisplayTableOutputValues(result.Data, AddHeaders);
            int recordCount = result.Data == null || result.Data.Rows == null ? 0 : result.Data.Rows.Count;
            OperationResult writeResult = _excelOutputService.WriteValuesToActiveCell(
                values,
                "Successfully wrote " + recordCount + " " + _displayTableName + " record(s) to Excel.",
                AddHeaders);

            StatusText = writeResult.Message;
            AnalysisExportDiagnostics.Log("Report export completed for " + _displayTableName + ": " + writeResult.Message);
            MessageBox.Show(
                writeResult.Message,
                WindowTitle,
                MessageBoxButton.OK,
                writeResult.IsSuccess ? MessageBoxImage.Information : MessageBoxImage.Warning);
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
                    AnalysisExportDiagnostics.Log("Failed to apply selected output unit for " + _displayTableName + ": " + unitScopeResult.Message);
                    ShowWarning(string.IsNullOrWhiteSpace(unitScopeResult.Message)
                        ? "Failed to set ETABS present units."
                        : unitScopeResult.Message);
                    return unitScopeResult;
                }

                AnalysisExportDiagnostics.Log("Selected output unit: " + SelectedUnitOption.Label + " for " + _displayTableName);
                return unitScopeResult;
            }
            catch (Exception ex)
            {
                AnalysisExportDiagnostics.Log("Failed to set ETABS present units for " + _displayTableName + ": " + ex.Message);
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
                    "Failed to restore ETABS present units after " + _displayTableName + ": " + unitScope.RestoreResult.Message);
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
            int headerOffset = addHeaders ? 2 : 0;
            var values = new object[rows.Count + headerOffset, 13];

            if (addHeaders)
            {
                values[0, 0] = "Base Reactions";

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
                    values[1, col] = headers[col];
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

        private object[,] CreateDisplayTableOutputValues(CSISapModelDisplayTableDTO table, bool addHeaders)
        {
            IReadOnlyList<string> fields = table == null || table.FieldKeys == null
                ? new List<string>()
                : table.FieldKeys;
            IReadOnlyList<object[]> rows = table == null || table.Rows == null
                ? new List<object[]>()
                : table.Rows;

            int columnCount = fields.Count > 0 ? fields.Count : 1;
            int rowOffset = addHeaders ? 2 : 0;
            int rowCount = rows.Count > 0 ? rows.Count + rowOffset : rowOffset + 1;
            object[,] values = new object[rowCount, columnCount];

            if (addHeaders)
            {
                values[0, 0] = _displayTableName;

                if (fields.Count > 0)
                {
                    for (int columnIndex = 0; columnIndex < fields.Count; columnIndex++)
                    {
                        string headerName = fields[columnIndex];
                        if (SelectedUnitOption != null)
                        {
                            headerName = ApplyUnitToHeader(headerName, _displayTableName, SelectedUnitOption);
                        }
                        values[1, columnIndex] = headerName;
                    }
                }
            }

            if (rows.Count == 0)
            {
                values[rowOffset, 0] = "No records found";
                return values;
            }

            for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++)
            {
                object[] row = rows[rowIndex];
                for (int columnIndex = 0; columnIndex < columnCount; columnIndex++)
                {
                    values[rowIndex + rowOffset, columnIndex] = row != null && columnIndex < row.Length
                        ? row[columnIndex]
                        : null;
                }
            }

            return values;
        }

        private bool IsBaseReactionsTable()
        {
            return string.Equals(_displayTableName, "Base Reactions", StringComparison.OrdinalIgnoreCase);
        }

        private static string FormatUnitHeader(string name, string unit)
        {
            return string.IsNullOrWhiteSpace(unit) ? name : $"{name} ({unit})";
        }

        private static string ApplyUnitToHeader(string headerName, string displayTableName, BaseReactionUnitOption unitOption)
        {
            if (string.IsNullOrWhiteSpace(headerName) || unitOption == null)
            {
                return headerName;
            }

            string clean = headerName.Trim();
            if (clean.Contains("(") && clean.Contains(")"))
            {
                return clean;
            }

            string upper = clean.ToUpperInvariant();

            if (displayTableName.IndexOf("Mass Summary", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                if (upper.Contains("MASS"))
                {
                    return FormatUnitHeader(clean, "kg");
                }

                if (upper.Contains("WEIGHT"))
                {
                    return FormatUnitHeader(clean, unitOption.ForceUnit);
                }

                if (upper == "X" || upper == "Y" || upper == "Z" ||
                    upper.Contains("COORD") ||
                    upper.Contains("CENTER OF MASS") ||
                    upper.Contains("CENTEROFMASS") ||
                    upper == "XCM" || upper == "YCM" || upper == "ZCM")
                {
                    return FormatUnitHeader(clean, unitOption.LengthUnit);
                }
            }

            // Displacements & Lengths
            if (upper == "U1" || upper == "U2" || upper == "U3" || 
                upper == "UX" || upper == "UY" || upper == "UZ" || 
                upper == "DISPLACEMENT X" || upper == "DISPLACEMENTY" || 
                upper == "DISPLACEMENTX" || upper == "DISPLACEMENT Y")
            {
                if (displayTableName.IndexOf("Velocit", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return FormatUnitHeader(clean, $"{unitOption.LengthUnit}/s");
                }
                if (displayTableName.IndexOf("Acceleration", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return FormatUnitHeader(clean, $"{unitOption.LengthUnit}/s²");
                }
                return FormatUnitHeader(clean, unitOption.LengthUnit);
            }

            // Rotations
            if (upper == "R1" || upper == "R2" || upper == "R3" ||
                upper == "RX" || upper == "RY" || upper == "RZ")
            {
                if (displayTableName.IndexOf("Velocit", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return FormatUnitHeader(clean, "rad/s");
                }
                if (displayTableName.IndexOf("Acceleration", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return FormatUnitHeader(clean, "rad/s²");
                }
                return FormatUnitHeader(clean, "rad");
            }

            // Forces
            if (upper == "FX" || upper == "FY" || upper == "FZ" || 
                upper == "F1" || upper == "F2" || upper == "F3" ||
                upper == "P" || upper == "V2" || upper == "V3" || 
                upper == "VX" || upper == "VY")
            {
                return FormatUnitHeader(clean, unitOption.ForceUnit);
            }

            // Shell Forces (Force per unit length)
            if (upper == "F11" || upper == "F22" || upper == "F12" || 
                upper == "V13" || upper == "V23" || 
                upper == "FMAX" || upper == "FMIN" || upper == "VMAX")
            {
                return FormatUnitHeader(clean, $"{unitOption.ForceUnit}/{unitOption.LengthUnit}");
            }

            // Moments / Torques
            if (upper == "MX" || upper == "MY" || upper == "MZ" || 
                upper == "M1" || upper == "M2" || upper == "M3" ||
                upper == "T" || upper == "TX" || upper == "TY" || upper == "TZ")
            {
                return FormatUnitHeader(clean, unitOption.MomentUnit);
            }

            // Shell Moments (Moment per unit length)
            if (upper == "M11" || upper == "M22" || upper == "M12" || 
                upper == "MMAX" || upper == "MMIN")
            {
                return FormatUnitHeader(clean, $"{unitOption.MomentUnit}/{unitOption.LengthUnit}");
            }

            // Stresses (Force per unit area)
            if (upper == "S11" || upper == "S22" || upper == "S12" || 
                upper == "SMAX" || upper == "SMIN" || upper == "SVM" ||
                upper == "S13" || upper == "S23" ||
                upper == "SMAXOUTER" || upper == "SMINOUTER" || upper == "SVMOUTER")
            {
                return FormatUnitHeader(clean, $"{unitOption.ForceUnit}/{unitOption.LengthUnit}²");
            }

            // Stiffness
            if (upper == "STIFFNESS X" || upper == "STIFFNESS Y" || 
                upper == "STIFFNESSX" || upper == "STIFFNESSY")
            {
                return FormatUnitHeader(clean, $"{unitOption.ForceUnit}/{unitOption.LengthUnit}");
            }

            return clean;
        }

        private bool EnsureEtabs()
        {
            if (!string.Equals(_csiConnectionService.ProductName, "ETABS", StringComparison.OrdinalIgnoreCase))
            {
                ShowWarning(WindowTitle + " is available from the ETABS Toolbox only.");
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
            catch (Exception ex)
            {
                AnalysisExportDiagnostics.Log("Failed to read active Excel cell: " + ex.Message);
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

            if (_workbookState.HasStoredValue)
            {
                AddHeaders = _workbookState.AddHeaders;
            }
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

        private void SelectCurrentEtabsUnitIfRequested()
        {
            if (!_profile.DefaultToCurrentEtabsUnit || !_profile.ShowUnitSelector)
            {
                return;
            }

            AnalysisExportDiagnostics.Log("Retrieving current ETABS unit for " + _displayTableName);
            try
            {
                OperationResult<int> result = _csiConnectionService.GetPresentUnits();
                if (!result.IsSuccess)
                {
                    AnalysisExportDiagnostics.Log("Failed to retrieve current ETABS unit for " + _displayTableName + ": " + result.Message);
                    return;
                }

                foreach (BaseReactionUnitOption unitOption in UnitOptions)
                {
                    if (unitOption.EtabsUnitsCode == result.Data)
                    {
                        SelectedUnitOption = unitOption;
                        AnalysisExportDiagnostics.Log("Retrieved current ETABS unit: " + unitOption.Label + " for " + _displayTableName);
                        return;
                    }
                }

                AnalysisExportDiagnostics.Log("Current ETABS unit code is not supported by popup list: " + result.Data);
            }
            catch (Exception ex)
            {
                AnalysisExportDiagnostics.Log("Failed to retrieve current ETABS unit for " + _displayTableName + ": " + ex.Message);
            }
        }

        private static string CreateStateKey(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return "ETABSTable";
            }

            var chars = new List<char>();
            foreach (char ch in value)
            {
                if (char.IsLetterOrDigit(ch))
                {
                    chars.Add(ch);
                }
            }

            return chars.Count == 0 ? "ETABSTable" : new string(chars.ToArray());
        }

        private BaseReactionUnitOption FindUnitOption(BaseReactionUnitOption requestedUnit)
        {
            if (requestedUnit == null)
            {
                return null;
            }

            foreach (BaseReactionUnitOption unitOption in UnitOptions)
            {
                if (unitOption.EtabsUnitsCode == requestedUnit.EtabsUnitsCode ||
                    string.Equals(unitOption.Label, requestedUnit.Label, StringComparison.OrdinalIgnoreCase))
                {
                    return unitOption;
                }
            }

            return null;
        }

        private void SaveWorkbookState()
        {
            if (!_isWorkbookStateLoaded)
            {
                return;
            }

            PostprocessingWorkbookStateStore.Save(_workbookStateKey, new PostprocessingWorkbookState
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

        private void ShowWarning(string message)
        {
            AnalysisExportDiagnostics.Log("Analysis export warning for " + _displayTableName + ": " + message);
            MessageBox.Show(
                string.IsNullOrWhiteSpace(message) ? "The operation could not be completed." : message,
                WindowTitle,
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
        }

        private void ShowError(string message)
        {
            AnalysisExportDiagnostics.Log("Analysis export error for " + _displayTableName + ": " + message);
            MessageBox.Show(
                string.IsNullOrWhiteSpace(message) ? "An unexpected error occurred." : message,
                WindowTitle,
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

        public CSISapModelPresentUnitSystemDTO ToPresentUnitSystemDto()
        {
            switch (EtabsUnitsCode)
            {
                case 1:
                    return Unit(1, 1, 2);
                case 4:
                    return Unit(2, 2, 2);
                case 6:
                    return Unit(4, 6, 2);
                case 9:
                    return Unit(3, 4, 2);
                default:
                    return Unit(4, 6, 2);
            }
        }

        private static CSISapModelPresentUnitSystemDTO Unit(int forceUnit, int lengthUnit, int temperatureUnit)
        {
            return new CSISapModelPresentUnitSystemDTO
            {
                ForceUnit = forceUnit,
                LengthUnit = lengthUnit,
                TemperatureUnit = temperatureUnit
            };
        }
    }
}
