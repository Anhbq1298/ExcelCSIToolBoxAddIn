using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using System.Windows.Threading;
using ExcelCSIToolBox.Application.Interfaces.Etabs;
using ExcelCSIToolBox.Application.Modelling.DropPanels;
using ExcelCSIToolBox.Core.Abstractions.CSI;
using ExcelCSIToolBox.Core.Common.Results;
using ExcelCSIToolBox.Core.Contracts.CSI;
using ExcelCSIToolBoxAddIn.AddIn;
using ExcelCSIToolBoxAddIn.AddIn.Diagnostics;
using ExcelCSIToolBoxAddIn.UI.Common.Commands;
using ExcelCSIToolBoxAddIn.UI.Views;

namespace ExcelCSIToolBoxAddIn.UI.ViewModels
{
    public enum ColumnDropDialogResult
    {
        None,
        RequestColumnSelection,
        CreateDrop,
        Close
    }

    public sealed class DropPanelViewModel : ViewModelBase
    {


        private readonly ICSISapModelConnectionService _connectionService;
        private readonly IDropPanelEtabsService _dropPanelService;
        private readonly DropPanelGeometryProcessor _geometryProcessor;
        private readonly DropPanelSettingsStore _settingsStore;
        private readonly DropPanelExcelLogExporter _logExporter;
        private readonly Func<Window> _getWindow;
        private readonly Action _closeWindow;
        private readonly RelayCommand _selectColumnsCommand;
        private readonly RelayCommand _applyCommand;
        private readonly RelayCommand _exportLogCommand;
        private readonly RelayCommand _closeCommand;
        private DropPanelOptions _options;
        private List<DropPanelLogEntry> _lastLogEntries;
        private List<string> _lastSkippedColumnMessages;
        private bool _isBusy;
        private bool _isModificationStarted;
        private bool _isModelAvailable;
        private bool _isModelLocked;
        private string _statusMessage;
        private string _validationMessage;
        private string _dropThicknessText;
        private string _dropWidthText;
        private string _dropLengthText;
        private string _rotationAngleText;
        private bool _isAlignWithColumnAxes;
        private bool _isSpecifiedRotation;
        private string _lengthUnit;
        private string _selectedMaterial;
        private int _ignoredNonColumnFrameCount;

        public ColumnDropDialogResult DialogResultState { get; set; }
        public ObservableCollection<EtabsUnitSystem> UnitSystems { get; private set; }

        private EtabsUnitSystem _selectedUnitSystem;
        public EtabsUnitSystem SelectedUnitSystem
        {
            get { return _selectedUnitSystem; }
            set
            {
                if (_selectedUnitSystem == value) return;
                var oldUnit = _selectedUnitSystem;
                _selectedUnitSystem = value;
                OnPropertyChanged();
                OnPropertyChanged("LengthUnit");
                OnPropertyChanged("PropertyNamePreview");
                ClearValidation();

                if (oldUnit != null && value != null)
                {
                    ConvertInputValues(oldUnit, value);
                }
                RaiseCanExecuteChanged();
            }
        }

        public DropPanelViewModel(
            ICSISapModelConnectionService connectionService,
            IDropPanelEtabsService dropPanelService,
            DropPanelSettingsStore settingsStore,
            DropPanelExcelLogExporter logExporter,
            Func<Window> getWindow,
            Action closeWindow)
        {
            _connectionService = connectionService ?? throw new ArgumentNullException(nameof(connectionService));
            _dropPanelService = dropPanelService ?? throw new ArgumentNullException(nameof(dropPanelService));
            _settingsStore = settingsStore ?? throw new ArgumentNullException(nameof(settingsStore));
            _logExporter = logExporter ?? throw new ArgumentNullException(nameof(logExporter));
            _getWindow = getWindow ?? delegate { return null; };
            _closeWindow = closeWindow ?? delegate { };
            _geometryProcessor = new DropPanelGeometryProcessor();
            _options = _settingsStore.Load();
            _lastLogEntries = new List<DropPanelLogEntry>();
            _lastSkippedColumnMessages = new List<string>();

            SelectedColumns = new ObservableCollection<DropPanelColumnInfo>();
            ConcreteMaterials = new ObservableCollection<string>();

            UnitSystems = new ObservableCollection<EtabsUnitSystem>
            {
                new EtabsUnitSystem("kN-m", 4, 6, 2, "kN", "kN-m", "m", 6),
                new EtabsUnitSystem("N-mm", 3, 4, 2, "N", "N-mm", "mm", 9),
                new EtabsUnitSystem("kip-ft", 2, 2, 2, "kip", "kip-ft", "ft", 4),
                new EtabsUnitSystem("lb-in", 1, 1, 2, "lb", "lb-in", "in", 1)
            };

            EtabsUnitSystem savedUnit = null;
            if (!string.IsNullOrEmpty(_options.LengthUnit))
            {
                savedUnit = UnitSystems.FirstOrDefault(u => string.Equals(u.LengthUnitText, _options.LengthUnit, StringComparison.OrdinalIgnoreCase));
            }

            var initialUnit = savedUnit;
            if (initialUnit == null)
            {
                var unitResult = _connectionService.GetPresentUnitSystem();
                if (unitResult.IsSuccess && unitResult.Data != null)
                {
                    initialUnit = UnitSystems.FirstOrDefault(u => u.Matches(unitResult.Data));
                }
            }
            _selectedUnitSystem = initialUnit ?? UnitSystems.FirstOrDefault(u => u.DisplayName == "N-mm") ?? UnitSystems.FirstOrDefault();

            double initialThickness = _options.DropThickness;
            double initialWidth = _options.DropSizeX > 0.0 ? _options.DropSizeX : 2.0;
            double initialLength = _options.DropSizeY > 0.0 ? _options.DropSizeY : 2.0;

            if (string.IsNullOrEmpty(_options.LengthUnit))
            {
                if (_selectedUnitSystem.LengthUnit != 6) // Meter
                {
                    var converter = new ExcelCSIToolBox.Application.Modelling.PileCaps.EtabsUnitConverter();
                    double mmThickness = initialThickness * 1000.0;
                    initialThickness = mmThickness / converter.GetMillimetersPerUnit(_selectedUnitSystem.LengthUnit);

                    double mmWidth = initialWidth * 1000.0;
                    initialWidth = mmWidth / converter.GetMillimetersPerUnit(_selectedUnitSystem.LengthUnit);

                    double mmLength = initialLength * 1000.0;
                    initialLength = mmLength / converter.GetMillimetersPerUnit(_selectedUnitSystem.LengthUnit);
                }
            }

            _dropThicknessText = initialThickness > 0.0 ? Format(initialThickness) : string.Empty;
            _dropWidthText = Format(initialWidth);
            _dropLengthText = Format(initialLength);
            _rotationAngleText = Format(_options.UserDefinedRotationAngle);
            _isSpecifiedRotation = _options.RotationMode == DropPanelRotationMode.UserDefinedAngle;
            _isAlignWithColumnAxes = !_isSpecifiedRotation;
            _selectedMaterial = _options.DropMaterial;
            _ignoredNonColumnFrameCount = 0;
            _lengthUnit = "-";
            _statusMessage = "Select one or more ETABS columns, then enter the drop definition.";
            _validationMessage = string.Empty;

            _selectColumnsCommand = new RelayCommand(SelectColumns, CanSelectColumns);
            _applyCommand = new RelayCommand(Apply, CanApply);
            _exportLogCommand = new RelayCommand(ExportLog, CanExportLog);
            _closeCommand = new RelayCommand(Close, CanClose);
            SelectColumnsCommand = _selectColumnsCommand;
            ApplyCommand = _applyCommand;
            ExportLogCommand = _exportLogCommand;
            CloseCommand = _closeCommand;

            RefreshModelInputs(true);
            AddInDiagnostics.Log("Column Drop Tool started.");
        }

        public ObservableCollection<DropPanelColumnInfo> SelectedColumns { get; private set; }

        public ObservableCollection<string> ConcreteMaterials { get; private set; }

        public ICommand SelectColumnsCommand { get; private set; }

        public ICommand ApplyCommand { get; private set; }

        public ICommand ExportLogCommand { get; private set; }

        public ICommand CloseCommand { get; private set; }

        public int SelectedColumnCount
        {
            get { return SelectedColumns.Count; }
        }

        public string SelectedColumnCountText
        {
            get { return "Selected Columns: " + SelectedColumnCount.ToString(CultureInfo.InvariantCulture); }
        }

        public string DropThicknessText
        {
            get { return _dropThicknessText; }
            set
            {
                if (_dropThicknessText == value) return;
                _dropThicknessText = value;
                OnPropertyChanged();
                OnPropertyChanged("PropertyNamePreview");
                ClearValidation();
                RaiseCanExecuteChanged();
            }
        }

        public string DropWidthText
        {
            get { return _dropWidthText; }
            set
            {
                if (_dropWidthText == value) return;
                _dropWidthText = value;
                OnPropertyChanged();
                ClearValidation();
                RaiseCanExecuteChanged();
            }
        }

        public string DropLengthText
        {
            get { return _dropLengthText; }
            set
            {
                if (_dropLengthText == value) return;
                _dropLengthText = value;
                OnPropertyChanged();
                ClearValidation();
                RaiseCanExecuteChanged();
            }
        }

        public bool IsAlignWithColumnAxes
        {
            get { return _isAlignWithColumnAxes; }
            set
            {
                if (_isAlignWithColumnAxes == value) return;
                _isAlignWithColumnAxes = value;
                if (value)
                {
                    _isSpecifiedRotation = false;
                    OnPropertyChanged("IsSpecifiedRotation");
                }

                OnPropertyChanged();
                ClearValidation();
                RaiseCanExecuteChanged();
            }
        }

        public bool IsSpecifiedRotation
        {
            get { return _isSpecifiedRotation; }
            set
            {
                if (_isSpecifiedRotation == value) return;
                _isSpecifiedRotation = value;
                if (value)
                {
                    _isAlignWithColumnAxes = false;
                    OnPropertyChanged("IsAlignWithColumnAxes");
                }

                OnPropertyChanged();
                ClearValidation();
                RaiseCanExecuteChanged();
            }
        }

        public string RotationAngleText
        {
            get { return _rotationAngleText; }
            set
            {
                if (_rotationAngleText == value) return;
                _rotationAngleText = value;
                OnPropertyChanged();
                ClearValidation();
                RaiseCanExecuteChanged();
            }
        }

        public string LengthUnit
        {
            get { return _lengthUnit; }
            private set { SetField(ref _lengthUnit, value, "LengthUnit"); }
        }

        public string SelectedMaterial
        {
            get { return _selectedMaterial; }
            set
            {
                if (string.Equals(_selectedMaterial, value, StringComparison.Ordinal)) return;
                _selectedMaterial = value;
                _options.DropMaterial = value;
                OnPropertyChanged();
                OnPropertyChanged("PropertyNamePreview");
                ClearValidation();
                RaiseCanExecuteChanged();
            }
        }        public string PropertyNamePreview
        {
            get
            {
                if (SelectedUnitSystem == null)
                {
                    return "-";
                }

                double thickness;
                if (!TryParsePositive(DropThicknessText, out thickness))
                {
                    return "-";
                }

                var converter = new ExcelCSIToolBox.Application.Modelling.PileCaps.EtabsUnitConverter();
                double thicknessInMm = Math.Round(thickness * converter.GetMillimetersPerUnit(SelectedUnitSystem.LengthUnit), 4);

                OperationResult<string> result = DropPanelPropertyNameBuilder.Build(thicknessInMm, SelectedMaterial);
                return result.IsSuccess ? result.Data : "-";
            }
        }

        public string StatusMessage
        {
            get { return _statusMessage; }
            set { SetField(ref _statusMessage, value, "StatusMessage"); }
        }

        public string ValidationMessage
        {
            get { return _validationMessage; }
            set { SetField(ref _validationMessage, value, "ValidationMessage"); }
        }

        public bool IsBusy
        {
            get { return _isBusy; }
            private set
            {
                if (_isBusy == value) return;
                _isBusy = value;
                OnPropertyChanged();
                RaiseCanExecuteChanged();
            }
        }

        public bool IsModificationStarted
        {
            get { return _isModificationStarted; }
            private set
            {
                if (_isModificationStarted == value) return;
                _isModificationStarted = value;
                OnPropertyChanged();
                RaiseCanExecuteChanged();
            }
        }

        public void RefreshModelInputs(bool showErrors)
        {
            if (IsBusy || IsModificationStarted)
            {
                return;
            }

            OperationResult<DropPanelModelContext> contextResult = _dropPanelService.GetModelContext();
            if (!contextResult.IsSuccess)
            {
                _isModelAvailable = false;
                _isModelLocked = false;
                LengthUnit = "-";
                if (showErrors)
                {
                    SetValidation(contextResult.Message);
                }

                RaiseCanExecuteChanged();
                return;
            }

            _isModelAvailable = true;
            _isModelLocked = contextResult.Data.IsLocked;
            LengthUnit = contextResult.Data.LengthUnit;

            OperationResult<IReadOnlyList<string>> materialResult = _dropPanelService.GetConcreteMaterialNames();
            if (!materialResult.IsSuccess)
            {
                ConcreteMaterials.Clear();
                SelectedMaterial = null;
                if (showErrors)
                {
                    SetValidation(materialResult.Message);
                }

                RaiseCanExecuteChanged();
                return;
            }

            string previousMaterial = SelectedMaterial;
            ConcreteMaterials.Clear();
            foreach (string materialName in materialResult.Data)
            {
                ConcreteMaterials.Add(materialName);
            }

            if (!string.IsNullOrWhiteSpace(previousMaterial) && ConcreteMaterials.Contains(previousMaterial))
            {
                SelectedMaterial = previousMaterial;
            }
            else
            {
                SelectedMaterial = ConcreteMaterials.FirstOrDefault();
            }

            RaiseCanExecuteChanged();
        }

        private void SelectColumns()
        {
            if (IsBusy)
            {
                return;
            }

            IsBusy = true;
            ClearValidation();
            Window toolWindow = _getWindow();
            try
            {
                OperationResult clearResult = _connectionService.ClearSelection();
                EnsureSuccess(clearResult);

                var selectionWindow = new InteractiveSelectionWindow(
                    "Select Columns",
                    "Select one or more column frame objects in ETABS, then click Use Columns.",
                    "Waiting for at least one frame object...",
                    "Select one or more frame objects.",
                    "Only frame objects can be used by the Column Drop Tool.",
                    "Frame",
                    () => _connectionService.GetSelectedObjectsFromActiveModel(),
                    true,
                    "Use Columns",
                    1,
                    "{0} frame object(s) selected. Click Use Columns to continue.",
                    "Select at least one column frame. Non-frame objects will be ignored.",
                    true);

                if (toolWindow != null && toolWindow.IsVisible)
                {
                    toolWindow.Hide();
                }

                selectionWindow.Owner = null;
                selectionWindow.WindowStartupLocation = WindowStartupLocation.CenterScreen;
                selectionWindow.ShowInTaskbar = false;
                selectionWindow.Topmost = true;
                selectionWindow.Loaded += delegate
                {
                    selectionWindow.Dispatcher.BeginInvoke(
                        new Action(ActivateConnectedCsiWindow),
                        DispatcherPriority.ApplicationIdle);
                };
                bool? selectionResult = selectionWindow.ShowDialog();
                if (selectionResult != true || selectionWindow.SelectedObjects == null)
                {
                    StatusMessage = "Column selection was cancelled. The previous selection was kept.";
                    return;
                }

                List<string> allFrameUniqueNames = selectionWindow.SelectedObjects
                    .Where(item => item != null &&
                           string.Equals(item.ObjectType, "Frame", StringComparison.OrdinalIgnoreCase))
                    .Select(item => item.UniqueName)
                    .Where(name => !string.IsNullOrWhiteSpace(name))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToList();

                if (allFrameUniqueNames.Count == 0)
                {
                    SetValidation("No valid column frame objects were found in the ETABS selection. Only frame objects with labels starting with 'C' are accepted.");
                    return;
                }

                var processResult = ProcessAndSetSelectedColumns(allFrameUniqueNames);
                if (!processResult.IsSuccess)
                {
                    SetValidation(processResult.Message);
                }
            }
            catch (Exception ex)
            {
                SetValidation(ex.Message);
                AddInDiagnostics.LogException("Column Drop interactive selection", ex);
            }
            finally
            {
                IsBusy = false;
                RestoreToolWindow(toolWindow);
            }
        }

        public OperationResult ProcessAndSetSelectedColumns(IReadOnlyList<string> frameUniqueNames)
        {
            if (frameUniqueNames == null || frameUniqueNames.Count == 0)
            {
                return OperationResult.Failure("No valid column frame objects were found in the ETABS selection. Only frame objects with labels starting with 'C' are accepted.");
            }

            OperationResult<IReadOnlyDictionary<string, string>> labelsResult =
                _dropPanelService.GetFrameLabels(frameUniqueNames);
            if (!labelsResult.IsSuccess)
            {
                return OperationResult.Failure(labelsResult.Message);
            }

            IReadOnlyDictionary<string, string> labelMap = labelsResult.Data;
            List<string> columnFrameNames = frameUniqueNames
                .Where(name =>
                {
                    string lbl;
                    return labelMap.TryGetValue(name, out lbl) &&
                           !string.IsNullOrEmpty(lbl) &&
                           lbl.StartsWith("C", StringComparison.OrdinalIgnoreCase);
                })
                .ToList();

            _ignoredNonColumnFrameCount = frameUniqueNames.Count - columnFrameNames.Count;

            if (columnFrameNames.Count == 0)
            {
                return OperationResult.Failure("No valid column frame objects were found in the ETABS selection. Only frame objects with labels starting with 'C' are accepted.");
            }

            OperationResult<IReadOnlyList<DropPanelColumnInfo>> columnsResult =
                _dropPanelService.ReadColumns(columnFrameNames, _options.VerticalRatioTolerance);
            if (!columnsResult.IsSuccess)
            {
                return OperationResult.Failure(columnsResult.Message);
            }

            List<DropPanelColumnInfo> validColumns = columnsResult.Data
                .Where(column => column != null && column.IsValid)
                .GroupBy(column => column.FrameName, StringComparer.OrdinalIgnoreCase)
                .Select(group => group.First())
                .ToList();

            _lastSkippedColumnMessages = columnsResult.Data
                .Where(column => column == null || !column.IsValid)
                .Select(column => column == null
                    ? "An unreadable frame selection was skipped."
                    : "'" + column.FrameName + "': " + column.ValidationMessage)
                .ToList();

            if (validColumns.Count == 0)
            {
                return OperationResult.Failure("No selected frame is a valid column. " +
                                              string.Join(" ", _lastSkippedColumnMessages));
            }

            ReplaceSelectedColumns(validColumns);
            StatusMessage = validColumns.Count.ToString(CultureInfo.InvariantCulture) +
                            " column(s) selected." + FormatSkippedColumns();
            ValidationMessage = string.Empty;
            return OperationResult.Success();
        }

        private void ConvertInputValues(EtabsUnitSystem fromUnit, EtabsUnitSystem toUnit)
        {
            if (fromUnit == null || toUnit == null || fromUnit.LengthUnit == toUnit.LengthUnit)
            {
                return;
            }

            var converter = new ExcelCSIToolBox.Application.Modelling.PileCaps.EtabsUnitConverter();
            
            double thickness;
            if (TryParsePositive(DropThicknessText, out thickness))
            {
                double mm = thickness * converter.GetMillimetersPerUnit(fromUnit.LengthUnit);
                double converted = mm / converter.GetMillimetersPerUnit(toUnit.LengthUnit);
                DropThicknessText = Format(converted);
            }

            double width;
            if (TryParsePositive(DropWidthText, out width))
            {
                double mm = width * converter.GetMillimetersPerUnit(fromUnit.LengthUnit);
                double converted = mm / converter.GetMillimetersPerUnit(toUnit.LengthUnit);
                DropWidthText = Format(converted);
            }

            double length;
            if (TryParsePositive(DropLengthText, out length))
            {
                double mm = length * converter.GetMillimetersPerUnit(fromUnit.LengthUnit);
                double converted = mm / converter.GetMillimetersPerUnit(toUnit.LengthUnit);
                DropLengthText = Format(converted);
            }
        }
        private async void Apply()
        {
            if (IsBusy)
            {
                return;
            }

            DropPanelOptions options;
            string validationMessage;
            if (!TryBuildOptions(out options, out validationMessage))
            {
                SetValidation(validationMessage);
                return;
            }

            IsBusy = true;
            ClearValidation();
            try
            {
                List<string> selectedNames = SelectedColumns
                    .Select(column => column.FrameName)
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToList();
                OperationResult<IReadOnlyList<DropPanelColumnInfo>> columnResult =
                    _dropPanelService.ReadColumns(selectedNames, options.VerticalRatioTolerance);
                EnsureSuccess(columnResult);
                List<DropPanelColumnInfo> refreshedColumns = columnResult.Data
                    .Where(column => column != null && column.IsValid)
                    .ToList();
                if (refreshedColumns.Count != selectedNames.Count)
                {
                    string invalidColumns = string.Join(" ", columnResult.Data
                        .Where(column => column == null || !column.IsValid)
                        .Select(column => column == null
                            ? "An unreadable selected column was found."
                            : "'" + column.FrameName + "': " + column.ValidationMessage));
                    throw new InvalidOperationException(
                        "One or more selected columns are no longer valid. " + invalidColumns);
                }

                ReplaceSelectedColumns(refreshedColumns);
                StatusMessage = "Reading connected slab shells and assignments...";
                OperationResult<IReadOnlyList<DropPanelRequest>> requestResult =
                    _geometryProcessor.BuildDropRequests(refreshedColumns, options);
                EnsureSuccess(requestResult);

                OperationResult<DropPanelPreparationSnapshot> snapshotResult =
                    _dropPanelService.PrepareSnapshot(refreshedColumns, requestResult.Data, options);
                EnsureSuccess(snapshotResult);

                StatusMessage = "Building drop boundaries and split regions...";
                OperationResult<DropPanelOperationPlan> planResult = await Task.Run(
                    () => _geometryProcessor.BuildPlan(
                        refreshedColumns,
                        snapshotResult.Data,
                        requestResult.Data,
                        options));
                EnsureSuccess(planResult);
                if (!planResult.Data.IsValid)
                {
                    throw new InvalidOperationException(
                        "Drop panel geometry is invalid:" + Environment.NewLine +
                        string.Join(Environment.NewLine, planResult.Data.ValidationMessages));
                }

                IsModificationStarted = true;
                StatusMessage = "Creating or reusing the drop property and updating ETABS shells...";
                OperationResult<DropPanelApplyResult> applyResult =
                    _dropPanelService.Apply(planResult.Data, options);
                EnsureSuccess(applyResult);

                _options = options;
                _options.DropProperty = applyResult.Data.DropPropertyName;
                _settingsStore.Save(_options);
                _lastLogEntries = applyResult.Data.LogEntries;
                StatusMessage = BuildSuccessSummary(applyResult.Data, options);
                AddInDiagnostics.Log(StatusMessage);
            }
            catch (Exception ex)
            {
                SetValidation(ex.Message);
                StatusMessage = "Column drop creation did not complete.";
                AddInDiagnostics.LogException("Column Drop apply", ex);
            }
            finally
            {
                IsModificationStarted = false;
                IsBusy = false;
            }
        }

        private void ExportLog()
        {
            if (IsBusy)
            {
                return;
            }

            try
            {
                OperationResult result = _logExporter.Export(_lastLogEntries);
                EnsureSuccess(result);
                StatusMessage = result.Message;
            }
            catch (Exception ex)
            {
                SetValidation(ex.Message);
            }
        }

        private void Close()
        {
            _closeWindow();
        }

        public bool TryCloseWindow()
        {
            if (IsBusy || IsModificationStarted)
            {
                return false;
            }

            try
            {
                DropPanelOptions options;
                string ignored;
                if (TryBuildOptions(out options, out ignored))
                {
                    _settingsStore.Save(options);
                }
            }
            catch (Exception ex)
            {
                AddInDiagnostics.LogException("Save Column Drop settings", ex);
            }

            return true;
        }

        private bool TryBuildOptions(out DropPanelOptions options, out string message)
        {
            options = CloneOptions(_options);
            message = string.Empty;

            var activeUnitResult = _connectionService.GetPresentUnitSystem();
            if (!activeUnitResult.IsSuccess || activeUnitResult.Data == null)
            {
                message = "The active ETABS model and its current length unit are unavailable.";
                return false;
            }
            int etabsLengthUnit = activeUnitResult.Data.LengthUnit;

            if (_isModelLocked)
            {
                message = "The ETABS model is locked. Unlock it before creating column drops.";
                return false;
            }

            double thickness;
            if (!TryParsePositive(DropThicknessText, out thickness))
            {
                message = "Drop thickness must be a numeric value greater than zero.";
                return false;
            }

            double width;
            if (!TryParsePositive(DropWidthText, out width))
            {
                message = "Drop width must be a numeric value greater than zero.";
                return false;
            }

            double length;
            if (!TryParsePositive(DropLengthText, out length))
            {
                message = "Drop length must be a numeric value greater than zero.";
                return false;
            }

            if (!IsAlignWithColumnAxes && !IsSpecifiedRotation)
            {
                message = "Select an orientation mode (Align with Column Axes or Specified Rotation).";
                return false;
            }

            double rotationAngle = 0.0;
            if (IsSpecifiedRotation && !TryParseDouble(RotationAngleText, out rotationAngle))
            {
                message = "Rotation angle must be a numeric value.";
                return false;
            }

            if (string.IsNullOrWhiteSpace(SelectedMaterial) || !ConcreteMaterials.Contains(SelectedMaterial))
            {
                message = "Select a concrete material that exists in the active ETABS model.";
                return false;
            }

            var converter = new ExcelCSIToolBox.Application.Modelling.PileCaps.EtabsUnitConverter();
            int selectedLengthUnit = SelectedUnitSystem.LengthUnit;
            double thicknessInMm = Math.Round(thickness * converter.GetMillimetersPerUnit(selectedLengthUnit), 4);

            OperationResult<string> propertyNameResult =
                DropPanelPropertyNameBuilder.Build(thicknessInMm, SelectedMaterial);
            if (!propertyNameResult.IsSuccess)
            {
                message = propertyNameResult.Message;
                return false;
            }

            double thicknessInEtabsUnit = thickness;
            double widthInEtabsUnit = width;
            double lengthInEtabsUnit = length;

            if (selectedLengthUnit != etabsLengthUnit)
            {
                double mmThickness = thickness * converter.GetMillimetersPerUnit(selectedLengthUnit);
                thicknessInEtabsUnit = mmThickness / converter.GetMillimetersPerUnit(etabsLengthUnit);

                double mmWidth = width * converter.GetMillimetersPerUnit(selectedLengthUnit);
                widthInEtabsUnit = mmWidth / converter.GetMillimetersPerUnit(etabsLengthUnit);

                double mmLength = length * converter.GetMillimetersPerUnit(selectedLengthUnit);
                lengthInEtabsUnit = mmLength / converter.GetMillimetersPerUnit(etabsLengthUnit);
            }

            options.DropThickness = thicknessInEtabsUnit;
            options.DropSizeX = widthInEtabsUnit;
            options.DropSizeY = lengthInEtabsUnit;
            options.DropMaterial = SelectedMaterial;
            options.LengthUnit = FormatLengthUnit(etabsLengthUnit);
            options.DropProperty = propertyNameResult.Data;
            options.RotationMode = IsSpecifiedRotation
                ? DropPanelRotationMode.UserDefinedAngle
                : DropPanelRotationMode.FollowColumnLocalAxis;
            options.UserDefinedRotationAngle = IsSpecifiedRotation ? rotationAngle : 0.0;
            return true;
        }

        private static string FormatLengthUnit(int lengthUnit)
        {
            switch (lengthUnit)
            {
                case 1: return "in";
                case 2: return "ft";
                case 3: return "micron";
                case 4: return "mm";
                case 5: return "cm";
                case 6: return "m";
                default: return "m";
            }
        }

        private static DropPanelOptions CloneOptions(DropPanelOptions source)
        {
            return new DropPanelOptions
            {
                DropProperty = source.DropProperty,
                DropThickness = source.DropThickness,
                DropMaterial = source.DropMaterial,
                LengthUnit = source.LengthUnit,
                DropSizeX = source.DropSizeX > 0.0 ? source.DropSizeX : 2.0,
                DropSizeY = source.DropSizeY > 0.0 ? source.DropSizeY : 2.0,
                RotationMode = source.RotationMode,
                UserDefinedRotationAngle = source.UserDefinedRotationAngle,
                GeometryTolerance = source.GeometryTolerance > 0.0 ? source.GeometryTolerance : 0.001,
                ElevationTolerance = source.ElevationTolerance >= 0.0 ? source.ElevationTolerance : 0.01,
                MinimumPolygonArea = source.MinimumPolygonArea > 0.0 ? source.MinimumPolygonArea : 0.001,
                VerticalRatioTolerance = source.VerticalRatioTolerance > 0.0 ? source.VerticalRatioTolerance : 4.0,
                PreserveDirectAreaLoads = true,
                PreserveShellUniformLoadSetAssignments = true,
                PreserveLocalAxes = true,
                PreserveLocal3Orientation = true,
                PreserveDiaphragm = true,
                PreserveMeshAssignments = true,
                PreserveAreaModifiers = true,
                PreserveGroupAssignments = true,
                PreservePierAndSpandrelLabels = true
            };
        }

        private void ReplaceSelectedColumns(IEnumerable<DropPanelColumnInfo> columns)
        {
            SelectedColumns.Clear();
            foreach (DropPanelColumnInfo column in columns)
            {
                SelectedColumns.Add(column);
            }

            OnPropertyChanged("SelectedColumnCount");
            OnPropertyChanged("SelectedColumnCountText");
            RaiseCanExecuteChanged();
        }

        private string BuildSuccessSummary(DropPanelApplyResult result, DropPanelOptions options)
        {
            string propertyAction = result.DropPropertyCreated ? "created" : "reused";
            string unitSystemName = SelectedUnitSystem != null ? SelectedUnitSystem.DisplayName : (result.LengthUnit ?? options.LengthUnit ?? LengthUnit);
            string orientationModeText = options.RotationMode == DropPanelRotationMode.FollowColumnLocalAxis
                ? "Align with Column Axes"
                : "Specified Rotation";

            var sb = new StringBuilder();
            sb.Append("Column drops created successfully.");
            sb.Append(" Columns processed: " + result.ProcessedColumnCount.ToString(CultureInfo.InvariantCulture));
            sb.Append("; drop objects: " + result.CreatedDropAreaCount.ToString(CultureInfo.InvariantCulture));
            sb.Append("; property: '" + result.DropPropertyName + "' (" + propertyAction + ")");
            sb.Append("; thickness: " + DropThicknessText + " (" + unitSystemName + ")");
            sb.Append("; width: " + DropWidthText + " (" + unitSystemName + ")");
            sb.Append("; length: " + DropLengthText + " (" + unitSystemName + ")");
            sb.Append("; material: " + result.MaterialName);
            sb.Append("; orientation: " + orientationModeText);
            if (options.RotationMode == DropPanelRotationMode.UserDefinedAngle)
            {
                sb.Append(" (Rotation Angle: " + RotationAngleText + " deg)");
            }
            sb.Append(".");
            sb.Append(FormatSkippedColumns());
            if (_ignoredNonColumnFrameCount > 0)
            {
                sb.Append(" " + _ignoredNonColumnFrameCount.ToString(CultureInfo.InvariantCulture) +
                          " non-column frame(s) were excluded by label filter.");
            }

            return sb.ToString();
        }

        private string FormatSkippedColumns()
        {
            return _lastSkippedColumnMessages.Count == 0
                ? string.Empty
                : " Skipped: " + string.Join(" ", _lastSkippedColumnMessages);
        }



        private bool CanSelectColumns()
        {
            return !IsBusy && !IsModificationStarted && _isModelAvailable;
        }

        private bool CanApply()
        {
            double thickness;
            double width;
            double length;
            double rotationAngle;
            return !IsBusy && !IsModificationStarted && _isModelAvailable && !_isModelLocked &&
                   SelectedColumns.Count > 0 &&
                   TryParsePositive(DropThicknessText, out thickness) &&
                   TryParsePositive(DropWidthText, out width) &&
                   TryParsePositive(DropLengthText, out length) &&
                   !string.IsNullOrWhiteSpace(SelectedMaterial) &&
                   ConcreteMaterials.Contains(SelectedMaterial) &&
                   (IsAlignWithColumnAxes || IsSpecifiedRotation) &&
                   (!IsSpecifiedRotation || TryParseDouble(RotationAngleText, out rotationAngle));
        }

        private bool CanExportLog()
        {
            return !IsBusy && _lastLogEntries.Count > 0;
        }

        private bool CanClose()
        {
            return !IsBusy && !IsModificationStarted;
        }

        private void RaiseCanExecuteChanged()
        {
            _selectColumnsCommand.RaiseCanExecuteChanged();
            _applyCommand.RaiseCanExecuteChanged();
            _exportLogCommand.RaiseCanExecuteChanged();
            _closeCommand.RaiseCanExecuteChanged();
        }

        private void SetValidation(string message)
        {
            ValidationMessage = message ?? string.Empty;
        }

        private void ClearValidation()
        {
            if (!string.IsNullOrEmpty(ValidationMessage))
            {
                ValidationMessage = string.Empty;
            }
        }

        private void SetField(ref string field, string value, string propertyName)
        {
            if (field == value) return;
            field = value;
            OnPropertyChanged(propertyName);
        }

        private static void EnsureSuccess(OperationResult result)
        {
            if (result == null || !result.IsSuccess)
            {
                throw new InvalidOperationException(
                    result == null ? "The operation returned no result." : result.Message);
            }
        }

        private static bool TryParsePositive(string text, out double value)
        {
            return (double.TryParse(text, NumberStyles.Float, CultureInfo.CurrentCulture, out value) ||
                    double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out value)) &&
                   !double.IsNaN(value) && !double.IsInfinity(value) && value > 0.0;
        }

        private static bool TryParseDouble(string text, out double value)
        {
            return (double.TryParse(text, NumberStyles.Float, CultureInfo.CurrentCulture, out value) ||
                    double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out value)) &&
                   !double.IsNaN(value) && !double.IsInfinity(value);
        }

        private static string Format(double value)
        {
            return value.ToString("0.################", CultureInfo.InvariantCulture);
        }

        private const int ShowWindowRestore = 9;

        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr windowHandle);

        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr windowHandle, int command);

        private void ActivateConnectedCsiWindow()
        {
            try
            {
                OperationResult<CSISapModelConnectionInfoDTO> connectionResult =
                    _connectionService.GetCurrentConnection();
                if (!connectionResult.IsSuccess ||
                    connectionResult.Data == null ||
                    !connectionResult.Data.ProcessId.HasValue)
                {
                    return;
                }

                Process process = Process.GetProcessById(connectionResult.Data.ProcessId.Value);
                if (process.MainWindowHandle == IntPtr.Zero)
                {
                    return;
                }

                ShowWindow(process.MainWindowHandle, ShowWindowRestore);
                SetForegroundWindow(process.MainWindowHandle);
            }
            catch
            {
                // Selection polling remains available if native window activation is unavailable.
            }
        }

        private static void RestoreToolWindow(Window toolWindow)
        {
            if (toolWindow == null)
            {
                return;
            }

            try
            {
                ModelessWpfWindowService.Show(toolWindow);
                toolWindow.Dispatcher.BeginInvoke(new Action(delegate
                {
                    try
                    {
                        IntPtr hwnd = new System.Windows.Interop.WindowInteropHelper(toolWindow).Handle;
                        if (hwnd != IntPtr.Zero)
                        {
                            ShowWindow(hwnd, ShowWindowRestore);
                            SetForegroundWindow(hwnd);
                        }
                    }
                    catch
                    {
                        // Fallback if interoperability helper fails
                    }

                    toolWindow.Activate();
                    toolWindow.Focus();
                    Keyboard.Focus(toolWindow);
                }));
            }
            catch (InvalidOperationException)
            {
                // The tool window was closed while the selection session was ending.
            }
        }
    }
}
