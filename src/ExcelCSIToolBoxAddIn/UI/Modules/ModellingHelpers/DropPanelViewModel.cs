using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using ExcelCSIToolBox.Application.Interfaces.Etabs;
using ExcelCSIToolBox.Application.Modelling.DropPanels;
using ExcelCSIToolBox.Core.Abstractions.CSI;
using ExcelCSIToolBox.Core.Common.Results;
using ExcelCSIToolBox.Core.Contracts.CSI;
using ExcelCSIToolBoxAddIn.AddIn.Diagnostics;
using ExcelCSIToolBoxAddIn.UI.Common.Commands;
using ExcelCSIToolBoxAddIn.UI.Views;

namespace ExcelCSIToolBoxAddIn.UI.ViewModels
{
    public sealed class DropPanelViewModel : ViewModelBase
    {
        private readonly ICSISapModelConnectionService _connectionService;
        private readonly IDropPanelEtabsService _dropPanelService;
        private readonly DropPanelGeometryProcessor _geometryProcessor;
        private readonly DropPanelSettingsStore _settingsStore;
        private readonly DropPanelExcelLogExporter _logExporter;
        private readonly Action _closeWindow;
        private readonly RelayCommand _attachCommand;
        private readonly RelayCommand _clearSelectionCommand;
        private readonly RelayCommand _readSelectedColumnsCommand;
        private readonly RelayCommand _previewCommand;
        private readonly RelayCommand _highlightCommand;
        private readonly RelayCommand _applyCommand;
        private readonly RelayCommand _rollbackCommand;
        private readonly RelayCommand _exportLogCommand;
        private readonly RelayCommand _cancelCommand;
        private readonly RelayCommand _closeCommand;
        private DropPanelOptions _options;
        private DropPanelOperationPlan _previewPlan;
        private List<DropPanelLogEntry> _lastLogEntries;
        private CancellationTokenSource _cancellationTokenSource;
        private string _connectionStatus;
        private string _etabsVersion;
        private string _modelFileName;
        private string _currentUnits;
        private string _modelLockStatus;
        private bool _isModelLocked;
        private bool _isBusy;
        private bool _isModificationStarted;
        private string _statusMessage;
        private string _dropSizeXText;
        private string _dropSizeYText;
        private string _userDefinedAngleText;
        private string _geometryToleranceText;
        private string _elevationToleranceText;
        private string _minimumPolygonAreaText;
        private string _selectedRotationMode;

        public DropPanelViewModel(
            ICSISapModelConnectionService connectionService,
            IDropPanelEtabsService dropPanelService,
            DropPanelSettingsStore settingsStore,
            DropPanelExcelLogExporter logExporter,
            Action closeWindow)
        {
            _connectionService = connectionService ?? throw new ArgumentNullException(nameof(connectionService));
            _dropPanelService = dropPanelService ?? throw new ArgumentNullException(nameof(dropPanelService));
            _settingsStore = settingsStore ?? throw new ArgumentNullException(nameof(settingsStore));
            _logExporter = logExporter ?? throw new ArgumentNullException(nameof(logExporter));
            _closeWindow = closeWindow ?? delegate { };
            _geometryProcessor = new DropPanelGeometryProcessor();
            _options = _settingsStore.Load();
            _lastLogEntries = new List<DropPanelLogEntry>();

            SelectedColumns = new ObservableCollection<DropPanelColumnInfo>();
            DropProperties = new ObservableCollection<string>();
            RotationModes = new ObservableCollection<string>
            {
                "Global X-Y",
                "Follow Column Local Axis",
                "User-Defined Angle"
            };

            _dropSizeXText = Format(_options.DropSizeX);
            _dropSizeYText = Format(_options.DropSizeY);
            _userDefinedAngleText = Format(_options.UserDefinedRotationAngle);
            _geometryToleranceText = Format(_options.GeometryTolerance);
            _elevationToleranceText = Format(_options.ElevationTolerance);
            _minimumPolygonAreaText = Format(_options.MinimumPolygonArea);
            _selectedRotationMode = ToDisplayName(_options.RotationMode);
            _connectionStatus = "Not attached";
            _etabsVersion = "-";
            _modelFileName = "-";
            _currentUnits = "-";
            _modelLockStatus = "Unknown";
            _statusMessage = "Attach to ETABS, select column frames, and read the current selection.";

            _attachCommand = new RelayCommand(AttachToEtabs, CanRunReadOperation);
            _clearSelectionCommand = new RelayCommand(ClearEtabsSelection, CanRunReadOperation);
            _readSelectedColumnsCommand = new RelayCommand(ReadSelectedColumns, CanRunReadOperation);
            _previewCommand = new RelayCommand(Preview, CanPreview);
            _highlightCommand = new RelayCommand(HighlightAffectedAreas, CanHighlight);
            _applyCommand = new RelayCommand(Apply, CanApply);
            _rollbackCommand = new RelayCommand(Rollback, CanRollback);
            _exportLogCommand = new RelayCommand(ExportLog, CanExportLog);
            _cancelCommand = new RelayCommand(Cancel, CanCancel);
            _closeCommand = new RelayCommand(Close, CanClose);

            AttachCommand = _attachCommand;
            ClearSelectionCommand = _clearSelectionCommand;
            ReadSelectedColumnsCommand = _readSelectedColumnsCommand;
            PreviewCommand = _previewCommand;
            HighlightAffectedAreasCommand = _highlightCommand;
            ApplyCommand = _applyCommand;
            RollbackCommand = _rollbackCommand;
            ExportLogCommand = _exportLogCommand;
            CancelCommand = _cancelCommand;
            CloseCommand = _closeCommand;

            RefreshContext(false);
            AddInDiagnostics.Log("Drop Panel tool started.");
        }

        public ObservableCollection<DropPanelColumnInfo> SelectedColumns { get; private set; }

        public ObservableCollection<string> DropProperties { get; private set; }

        public ObservableCollection<string> RotationModes { get; private set; }

        public ICommand AttachCommand { get; private set; }

        public ICommand ClearSelectionCommand { get; private set; }

        public ICommand ReadSelectedColumnsCommand { get; private set; }

        public ICommand PreviewCommand { get; private set; }

        public ICommand HighlightAffectedAreasCommand { get; private set; }

        public ICommand ApplyCommand { get; private set; }

        public ICommand RollbackCommand { get; private set; }

        public ICommand ExportLogCommand { get; private set; }

        public ICommand CancelCommand { get; private set; }

        public ICommand CloseCommand { get; private set; }

        public string ConnectionStatus
        {
            get { return _connectionStatus; }
            private set { SetField(ref _connectionStatus, value, "ConnectionStatus"); }
        }

        public string EtabsVersion
        {
            get { return _etabsVersion; }
            private set { SetField(ref _etabsVersion, value, "EtabsVersion"); }
        }

        public string ModelFileName
        {
            get { return _modelFileName; }
            private set { SetField(ref _modelFileName, value, "ModelFileName"); }
        }

        public string CurrentUnits
        {
            get { return _currentUnits; }
            private set { SetField(ref _currentUnits, value, "CurrentUnits"); }
        }

        public string ModelLockStatus
        {
            get { return _modelLockStatus; }
            private set { SetField(ref _modelLockStatus, value, "ModelLockStatus"); }
        }

        public string StatusMessage
        {
            get { return _statusMessage; }
            private set { SetField(ref _statusMessage, value, "StatusMessage"); }
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

        public DropPanelOperationPlan PreviewPlan
        {
            get { return _previewPlan; }
            private set
            {
                if (ReferenceEquals(_previewPlan, value)) return;
                _previewPlan = value;
                OnPropertyChanged();
                OnPropertyChanged("PreviewSummary");
                OnPropertyChanged("ValidationSummary");
                RaiseCanExecuteChanged();
            }
        }

        public string PreviewSummary
        {
            get
            {
                return PreviewPlan == null
                    ? "No preview is available."
                    : PreviewPlan.SourceAreas.Count.ToString(CultureInfo.InvariantCulture) + " source area(s), " +
                      PreviewPlan.Regions.Count(region => region.IsDrop).ToString(CultureInfo.InvariantCulture) + " drop region(s), " +
                      PreviewPlan.Regions.Count(region => !region.IsDrop).ToString(CultureInfo.InvariantCulture) + " normal region(s).";
            }
        }

        public string ValidationSummary
        {
            get
            {
                return PreviewPlan == null || PreviewPlan.ValidationMessages.Count == 0
                    ? string.Empty
                    : string.Join(Environment.NewLine, PreviewPlan.ValidationMessages);
            }
        }

        public string DropProperty
        {
            get { return _options.DropProperty; }
            set
            {
                if (string.Equals(_options.DropProperty, value, StringComparison.Ordinal)) return;
                _options.DropProperty = value;
                OnPropertyChanged();
                InvalidatePreview();
            }
        }

        public string DropSizeXText
        {
            get { return _dropSizeXText; }
            set { SetOptionText(ref _dropSizeXText, value, "DropSizeXText"); }
        }

        public string DropSizeYText
        {
            get { return _dropSizeYText; }
            set { SetOptionText(ref _dropSizeYText, value, "DropSizeYText"); }
        }

        public string SelectedRotationMode
        {
            get { return _selectedRotationMode; }
            set
            {
                if (_selectedRotationMode == value) return;
                _selectedRotationMode = value;
                OnPropertyChanged();
                OnPropertyChanged("IsUserDefinedRotation");
                InvalidatePreview();
            }
        }

        public bool IsUserDefinedRotation
        {
            get { return string.Equals(SelectedRotationMode, "User-Defined Angle", StringComparison.Ordinal); }
        }

        public string UserDefinedAngleText
        {
            get { return _userDefinedAngleText; }
            set { SetOptionText(ref _userDefinedAngleText, value, "UserDefinedAngleText"); }
        }

        public string GeometryToleranceText
        {
            get { return _geometryToleranceText; }
            set { SetOptionText(ref _geometryToleranceText, value, "GeometryToleranceText"); }
        }

        public string ElevationToleranceText
        {
            get { return _elevationToleranceText; }
            set { SetOptionText(ref _elevationToleranceText, value, "ElevationToleranceText"); }
        }

        public string MinimumPolygonAreaText
        {
            get { return _minimumPolygonAreaText; }
            set { SetOptionText(ref _minimumPolygonAreaText, value, "MinimumPolygonAreaText"); }
        }

        public bool PreserveDirectAreaLoads
        {
            get { return _options.PreserveDirectAreaLoads; }
            set { SetOption(_options.PreserveDirectAreaLoads, value, "PreserveDirectAreaLoads", newValue => _options.PreserveDirectAreaLoads = newValue); }
        }

        public bool PreserveShellUniformLoadSetAssignments
        {
            get { return _options.PreserveShellUniformLoadSetAssignments; }
            set { SetOption(_options.PreserveShellUniformLoadSetAssignments, value, "PreserveShellUniformLoadSetAssignments", newValue => _options.PreserveShellUniformLoadSetAssignments = newValue); }
        }

        public bool PreserveLocalAxes
        {
            get { return _options.PreserveLocalAxes; }
            set { SetOption(_options.PreserveLocalAxes, value, "PreserveLocalAxes", newValue => _options.PreserveLocalAxes = newValue); }
        }

        public bool PreserveLocal3Orientation
        {
            get { return _options.PreserveLocal3Orientation; }
            set { SetOption(_options.PreserveLocal3Orientation, value, "PreserveLocal3Orientation", newValue => _options.PreserveLocal3Orientation = newValue); }
        }

        public bool PreserveDiaphragm
        {
            get { return _options.PreserveDiaphragm; }
            set { SetOption(_options.PreserveDiaphragm, value, "PreserveDiaphragm", newValue => _options.PreserveDiaphragm = newValue); }
        }

        public bool PreserveMeshAssignments
        {
            get { return _options.PreserveMeshAssignments; }
            set { SetOption(_options.PreserveMeshAssignments, value, "PreserveMeshAssignments", newValue => _options.PreserveMeshAssignments = newValue); }
        }

        public bool PreserveAreaModifiers
        {
            get { return _options.PreserveAreaModifiers; }
            set { SetOption(_options.PreserveAreaModifiers, value, "PreserveAreaModifiers", newValue => _options.PreserveAreaModifiers = newValue); }
        }

        public bool PreserveGroupAssignments
        {
            get { return _options.PreserveGroupAssignments; }
            set { SetOption(_options.PreserveGroupAssignments, value, "PreserveGroupAssignments", newValue => _options.PreserveGroupAssignments = newValue); }
        }

        public bool PreservePierAndSpandrelLabels
        {
            get { return _options.PreservePierAndSpandrelLabels; }
            set { SetOption(_options.PreservePierAndSpandrelLabels, value, "PreservePierAndSpandrelLabels", newValue => _options.PreservePierAndSpandrelLabels = newValue); }
        }

        public bool SaveEtabsBackupBeforeApply
        {
            get { return _options.SaveEtabsBackupBeforeApply; }
            set { SetOption(_options.SaveEtabsBackupBeforeApply, value, "SaveEtabsBackupBeforeApply", newValue => _options.SaveEtabsBackupBeforeApply = newValue); }
        }

        public bool MergeAdjacentRegionsOnlyWhenAssignmentSignaturesMatch
        {
            get { return _options.MergeAdjacentRegionsOnlyWhenAssignmentSignaturesMatch; }
            set { SetOption(_options.MergeAdjacentRegionsOnlyWhenAssignmentSignaturesMatch, value, "MergeAdjacentRegionsOnlyWhenAssignmentSignaturesMatch", newValue => _options.MergeAdjacentRegionsOnlyWhenAssignmentSignaturesMatch = newValue); }
        }

        public bool VerifyAssignmentsAfterApply
        {
            get { return _options.VerifyAssignmentsAfterApply; }
            set { SetOption(_options.VerifyAssignmentsAfterApply, value, "VerifyAssignmentsAfterApply", newValue => _options.VerifyAssignmentsAfterApply = newValue); }
        }

        private void AttachToEtabs()
        {
            RunUiAction("Attach to ETABS", delegate
            {
                OperationResult<CSISapModelConnectionInfoDTO> attachResult = _connectionService.TryAttachToRunningInstance();
                if (!attachResult.IsSuccess)
                {
                    throw new InvalidOperationException(attachResult.Message);
                }

                RefreshContext(true);
                StatusMessage = "Attached to ETABS. Select column frames in ETABS and click Read Selected Columns.";
                AddInDiagnostics.Log("Drop Panel attached to ETABS.");
            });
        }

        private void ClearEtabsSelection()
        {
            RunUiAction("Clear ETABS Selection", delegate
            {
                OperationResult result = _dropPanelService.HighlightAreas(new string[0]);
                EnsureSuccess(result);
                SelectedColumns.Clear();
                PreviewPlan = null;
                StatusMessage = result.Message;
            });
        }

        private void ReadSelectedColumns()
        {
            RunUiAction("Read Selected Columns", delegate
            {
                OperationResult<IReadOnlyList<DropPanelColumnInfo>> result =
                    _dropPanelService.ReadSelectedColumns(_options.VerticalRatioTolerance);
                EnsureSuccess(result);
                SelectedColumns.Clear();
                foreach (DropPanelColumnInfo column in result.Data)
                {
                    SelectedColumns.Add(column);
                }

                PreviewPlan = null;
                int validCount = SelectedColumns.Count(column => column.IsValid);
                int invalidCount = SelectedColumns.Count - validCount;
                StatusMessage = "Read " + SelectedColumns.Count.ToString(CultureInfo.InvariantCulture) +
                                " selected object(s): " + validCount.ToString(CultureInfo.InvariantCulture) +
                                " valid column(s), " + invalidCount.ToString(CultureInfo.InvariantCulture) + " rejected.";
                AddInDiagnostics.Log("Drop Panel selected column count: " + validCount.ToString(CultureInfo.InvariantCulture) + ".");
            });
        }

        private async void Preview()
        {
            if (IsBusy)
            {
                return;
            }

            DropPanelOptions options;
            string validationMessage;
            if (!TryBuildOptions(out options, out validationMessage))
            {
                ShowWarning(validationMessage);
                return;
            }

            IsBusy = true;
            StatusMessage = "Reading ETABS slab geometry and assignments...";
            _cancellationTokenSource = new CancellationTokenSource();
            CancellationToken cancellationToken = _cancellationTokenSource.Token;
            RaiseCanExecuteChanged();
            try
            {
                IReadOnlyList<DropPanelColumnInfo> columnSnapshot = SelectedColumns.ToList();
                OperationResult<IReadOnlyList<DropPanelRequest>> requestResult =
                    _geometryProcessor.BuildDropRequests(columnSnapshot, options);
                EnsureSuccess(requestResult);
                cancellationToken.ThrowIfCancellationRequested();

                OperationResult<DropPanelPreparationSnapshot> snapshotResult =
                    _dropPanelService.PrepareSnapshot(columnSnapshot, requestResult.Data, options);
                EnsureSuccess(snapshotResult);
                AddInDiagnostics.Log("Drop Panel affected source area candidates: " + snapshotResult.Data.Areas.Count.ToString(CultureInfo.InvariantCulture) + ".");
                cancellationToken.ThrowIfCancellationRequested();

                StatusMessage = "Processing batch geometry...";
                OperationResult<DropPanelOperationPlan> planResult = await Task.Run(
                    () => _geometryProcessor.BuildPlan(
                        columnSnapshot,
                        snapshotResult.Data,
                        requestResult.Data,
                        options),
                    cancellationToken);
                cancellationToken.ThrowIfCancellationRequested();
                EnsureSuccess(planResult);
                PreviewPlan = planResult.Data;
                _options = options;
                _settingsStore.Save(_options);
                if (PreviewPlan.IsValid)
                {
                    StatusMessage = "Preview ready: " + PreviewSummary;
                    AddInDiagnostics.Log("Drop Panel geometry validation passed. " + PreviewSummary);
                }
                else
                {
                    StatusMessage = "Preview contains validation errors. Apply is disabled.";
                    AddInDiagnostics.Log("Drop Panel geometry validation failed: " + ValidationSummary.Replace(Environment.NewLine, " | "));
                }
            }
            catch (OperationCanceledException)
            {
                StatusMessage = "Preview cancelled before ETABS model modification began.";
                AddInDiagnostics.Log("Drop Panel preview cancelled.");
            }
            catch (Exception ex)
            {
                PreviewPlan = null;
                StatusMessage = ex.Message;
                AddInDiagnostics.LogException("Drop Panel preview", ex);
                ShowError(ex.Message);
            }
            finally
            {
                _cancellationTokenSource.Dispose();
                _cancellationTokenSource = null;
                IsBusy = false;
            }
        }

        private void HighlightAffectedAreas()
        {
            RunUiAction("Highlight Affected Areas", delegate
            {
                IReadOnlyList<string> areaNames = PreviewPlan.SourceAreas
                    .Select(area => area.AreaName)
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToList();
                OperationResult result = _dropPanelService.HighlightAreas(areaNames);
                EnsureSuccess(result);
                StatusMessage = result.Message;
            });
        }

        private void Apply()
        {
            string backupNotice = SaveEtabsBackupBeforeApply
                ? "A model backup will be created before source areas are deleted."
                : "Backup creation is disabled. Automatic rollback will not be available if ETABS rejects a change.";
            if (MessageBox.Show(
                    "Apply the validated Drop Panel plan to ETABS? " + backupNotice,
                    "Apply Drop Panels",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Warning) != MessageBoxResult.Yes)
            {
                return;
            }

            DropPanelOptions options;
            string validationMessage;
            if (!TryBuildOptions(out options, out validationMessage))
            {
                ShowWarning(validationMessage);
                return;
            }

            IsBusy = true;
            IsModificationStarted = true;
            StatusMessage = "Applying the validated batch to ETABS. Cancellation is disabled after modification begins...";
            try
            {
                AddInDiagnostics.Log(
                    "Drop Panel apply started. Source areas: " +
                    string.Join(", ", PreviewPlan.SourceAreas.Select(area => area.AreaName).Distinct(StringComparer.OrdinalIgnoreCase)) +
                    "; backup enabled: " + options.SaveEtabsBackupBeforeApply + ".");
                OperationResult<DropPanelApplyResult> result = _dropPanelService.Apply(PreviewPlan, options);
                EnsureSuccess(result);
                _lastLogEntries = result.Data.LogEntries;
                OnPropertyChanged("PreviewSummary");
                if (result.Data.VerificationPassed)
                {
                    StatusMessage = result.Message + " Backup: " + (result.Data.BackupFilePath ?? "not created");
                    MessageBox.Show(StatusMessage, "Drop Panel", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                else
                {
                    StatusMessage = result.Message;
                    IReadOnlyList<string> failedAreas = result.Data.VerificationIssues
                        .Select(issue => issue.NewAreaName)
                        .Where(name => !string.IsNullOrWhiteSpace(name))
                        .Distinct(StringComparer.OrdinalIgnoreCase)
                        .ToList();
                    if (failedAreas.Count > 0)
                    {
                        OperationResult highlightResult = _dropPanelService.HighlightAreas(failedAreas);
                        AddInDiagnostics.Log(
                            highlightResult.IsSuccess
                                ? "Drop Panel highlighted failed ETABS areas: " + string.Join(", ", failedAreas) + "."
                                : "Drop Panel could not highlight failed ETABS areas: " + highlightResult.Message);
                    }

                    MessageBox.Show(
                        result.Message + Environment.NewLine + Environment.NewLine +
                        result.Data.VerificationIssues.Count.ToString(CultureInfo.InvariantCulture) + " verification issue(s) were recorded.",
                        "Drop Panel Verification Failed",
                        MessageBoxButton.OK,
                        MessageBoxImage.Error);
                }

                AddInDiagnostics.Log("Drop Panel created area count: " + result.Data.CreatedAreaNames.Count.ToString(CultureInfo.InvariantCulture) +
                                     "; verification passed: " + result.Data.VerificationPassed + ".");
                AddInDiagnostics.Log("Drop Panel backup path: " +
                                     (string.IsNullOrWhiteSpace(result.Data.BackupFilePath) ? "not created" : result.Data.BackupFilePath) + ".");
                AddInDiagnostics.Log("Drop Panel created areas: " + string.Join(", ", result.Data.CreatedAreaNames) + ".");
                AddInDiagnostics.Log(
                    "Drop Panel deleted source areas: " +
                    string.Join(", ", PreviewPlan.SourceAreas.Select(area => area.AreaName).Distinct(StringComparer.OrdinalIgnoreCase)) + ".");
                AddInDiagnostics.Log("Drop Panel assignment restoration and verification completed: " + result.Data.VerificationPassed + ".");
                foreach (DropPanelVerificationIssue issue in result.Data.VerificationIssues)
                {
                    AddInDiagnostics.Log(
                        "Drop Panel verification issue for '" + issue.NewAreaName + "' / " + issue.AssignmentType +
                        ": expected '" + issue.ExpectedValue + "', actual '" + issue.ActualValue + "'. " + issue.ErrorMessage);
                }
                RefreshContext(false);
            }
            catch (Exception ex)
            {
                StatusMessage = ex.Message;
                AddInDiagnostics.LogException("Drop Panel apply", ex);
                ShowError(ex.Message);
            }
            finally
            {
                IsModificationStarted = false;
                IsBusy = false;
            }
        }

        private void Rollback()
        {
            if (MessageBox.Show(
                    "Restore the ETABS model from the last Drop Panel backup?",
                    "Rollback Drop Panel Operation",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Warning) != MessageBoxResult.Yes)
            {
                return;
            }

            RunUiAction("Rollback", delegate
            {
                OperationResult result = _dropPanelService.Rollback();
                EnsureSuccess(result);
                PreviewPlan = null;
                StatusMessage = result.Message;
                AddInDiagnostics.Log("Drop Panel rollback completed.");
                RefreshContext(false);
            });
        }

        private void ExportLog()
        {
            RunUiAction("Export Log", delegate
            {
                OperationResult result = _logExporter.Export(_lastLogEntries);
                EnsureSuccess(result);
                StatusMessage = result.Message;
            });
        }

        private void Cancel()
        {
            if (_cancellationTokenSource != null && !IsModificationStarted)
            {
                _cancellationTokenSource.Cancel();
            }
        }

        private void Close()
        {
            _closeWindow();
        }

        public bool TryCloseWindow()
        {
            if (IsModificationStarted)
            {
                return false;
            }

            if (_cancellationTokenSource != null)
            {
                _cancellationTokenSource.Cancel();
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
                AddInDiagnostics.LogException("Save Drop Panel settings", ex);
            }

            return true;
        }

        private void RefreshContext(bool showErrors)
        {
            OperationResult<DropPanelModelContext> contextResult = _dropPanelService.GetModelContext();
            if (!contextResult.IsSuccess)
            {
                ConnectionStatus = "Not attached";
                EtabsVersion = "-";
                ModelFileName = "-";
                CurrentUnits = "-";
                ModelLockStatus = "Unknown";
                _isModelLocked = false;
                if (showErrors)
                {
                    ShowWarning(contextResult.Message);
                }
                RaiseCanExecuteChanged();
                return;
            }

            DropPanelModelContext context = contextResult.Data;
            ConnectionStatus = "Attached";
            EtabsVersion = context.Version;
            ModelFileName = string.IsNullOrWhiteSpace(context.ModelFileName) ? "Unsaved model" : context.ModelFileName;
            CurrentUnits = context.PresentUnits;
            _isModelLocked = context.IsLocked;
            ModelLockStatus = context.IsLocked ? "Locked" : "Unlocked";

            OperationResult<IReadOnlyList<string>> propertyResult = _dropPanelService.GetDropPropertyNames();
            if (propertyResult.IsSuccess)
            {
                string selected = _options.DropProperty;
                DropProperties.Clear();
                foreach (string property in propertyResult.Data)
                {
                    DropProperties.Add(property);
                }

                if (!string.IsNullOrWhiteSpace(selected) && DropProperties.Contains(selected))
                {
                    DropProperty = selected;
                }
                else if (DropProperties.Count > 0)
                {
                    DropProperty = DropProperties[0];
                }
            }

            RaiseCanExecuteChanged();
        }

        private bool TryBuildOptions(out DropPanelOptions options, out string message)
        {
            options = CloneOptions(_options);
            message = string.Empty;
            double value;
            if (!TryParsePositive(DropSizeXText, out value))
            {
                message = "Drop Size X must be a positive number.";
                return false;
            }
            options.DropSizeX = value;

            if (!TryParsePositive(DropSizeYText, out value))
            {
                message = "Drop Size Y must be a positive number.";
                return false;
            }
            options.DropSizeY = value;

            if (!TryParseNumber(UserDefinedAngleText, out value))
            {
                message = "User-Defined Rotation Angle must be a number.";
                return false;
            }
            options.UserDefinedRotationAngle = value;

            if (!TryParsePositive(GeometryToleranceText, out value))
            {
                message = "Geometry Tolerance must be a positive number.";
                return false;
            }
            options.GeometryTolerance = value;

            if (!TryParseNumber(ElevationToleranceText, out value) || value < 0.0)
            {
                message = "Elevation Tolerance must be zero or greater.";
                return false;
            }
            options.ElevationTolerance = value;

            if (!TryParsePositive(MinimumPolygonAreaText, out value))
            {
                message = "Minimum Polygon Area must be a positive number.";
                return false;
            }
            options.MinimumPolygonArea = value;
            options.RotationMode = FromDisplayName(SelectedRotationMode);
            options.DropProperty = DropProperty;
            if (string.IsNullOrWhiteSpace(options.DropProperty))
            {
                message = "Select a Drop Property.";
                return false;
            }

            return true;
        }

        private static DropPanelOptions CloneOptions(DropPanelOptions source)
        {
            return new DropPanelOptions
            {
                DropProperty = source.DropProperty,
                DropSizeX = source.DropSizeX,
                DropSizeY = source.DropSizeY,
                RotationMode = source.RotationMode,
                UserDefinedRotationAngle = source.UserDefinedRotationAngle,
                GeometryTolerance = source.GeometryTolerance,
                ElevationTolerance = source.ElevationTolerance,
                MinimumPolygonArea = source.MinimumPolygonArea,
                VerticalRatioTolerance = source.VerticalRatioTolerance,
                PreserveDirectAreaLoads = source.PreserveDirectAreaLoads,
                PreserveShellUniformLoadSetAssignments = source.PreserveShellUniformLoadSetAssignments,
                PreserveLocalAxes = source.PreserveLocalAxes,
                PreserveLocal3Orientation = source.PreserveLocal3Orientation,
                PreserveDiaphragm = source.PreserveDiaphragm,
                PreserveMeshAssignments = source.PreserveMeshAssignments,
                PreserveAreaModifiers = source.PreserveAreaModifiers,
                PreserveGroupAssignments = source.PreserveGroupAssignments,
                PreservePierAndSpandrelLabels = source.PreservePierAndSpandrelLabels,
                SaveEtabsBackupBeforeApply = source.SaveEtabsBackupBeforeApply,
                MergeAdjacentRegionsOnlyWhenAssignmentSignaturesMatch = source.MergeAdjacentRegionsOnlyWhenAssignmentSignaturesMatch,
                VerifyAssignmentsAfterApply = source.VerifyAssignmentsAfterApply
            };
        }

        private bool CanRunReadOperation()
        {
            return !IsBusy && !IsModificationStarted;
        }

        private bool CanPreview()
        {
            return !IsBusy && !IsModificationStarted && SelectedColumns.Any(column => column.IsValid) && !_isModelLocked;
        }

        private bool CanHighlight()
        {
            return !IsBusy && PreviewPlan != null && PreviewPlan.SourceAreas.Count > 0;
        }

        private bool CanApply()
        {
            return !IsBusy && !IsModificationStarted && !_isModelLocked && PreviewPlan != null && PreviewPlan.IsValid;
        }

        private bool CanRollback()
        {
            return !IsBusy && !IsModificationStarted && _dropPanelService.IsRollbackAvailable;
        }

        private bool CanExportLog()
        {
            return !IsBusy && _lastLogEntries.Count > 0;
        }

        private bool CanCancel()
        {
            return IsBusy && !IsModificationStarted && _cancellationTokenSource != null;
        }

        private bool CanClose()
        {
            return !IsModificationStarted;
        }

        private void RaiseCanExecuteChanged()
        {
            _attachCommand.RaiseCanExecuteChanged();
            _clearSelectionCommand.RaiseCanExecuteChanged();
            _readSelectedColumnsCommand.RaiseCanExecuteChanged();
            _previewCommand.RaiseCanExecuteChanged();
            _highlightCommand.RaiseCanExecuteChanged();
            _applyCommand.RaiseCanExecuteChanged();
            _rollbackCommand.RaiseCanExecuteChanged();
            _exportLogCommand.RaiseCanExecuteChanged();
            _cancelCommand.RaiseCanExecuteChanged();
            _closeCommand.RaiseCanExecuteChanged();
        }

        private void RunUiAction(string context, Action action)
        {
            if (IsBusy)
            {
                return;
            }

            IsBusy = true;
            try
            {
                action();
            }
            catch (Exception ex)
            {
                StatusMessage = ex.Message;
                AddInDiagnostics.LogException("Drop Panel " + context, ex);
                ShowError(ex.Message);
            }
            finally
            {
                IsBusy = false;
            }
        }

        private void InvalidatePreview()
        {
            PreviewPlan = null;
            RaiseCanExecuteChanged();
        }

        private void SetOptionText(ref string field, string value, string propertyName)
        {
            if (field == value) return;
            field = value;
            OnPropertyChanged(propertyName);
            InvalidatePreview();
        }

        private void SetOption(bool currentValue, bool value, string propertyName, Action<bool> assign)
        {
            if (currentValue == value) return;
            assign(value);
            OnPropertyChanged(propertyName);
            InvalidatePreview();
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
                throw new InvalidOperationException(result == null ? "The operation returned no result." : result.Message);
            }
        }

        private static bool TryParsePositive(string text, out double value)
        {
            return TryParseNumber(text, out value) && value > 0.0;
        }

        private static bool TryParseNumber(string text, out double value)
        {
            return double.TryParse(text, NumberStyles.Float, CultureInfo.CurrentCulture, out value) ||
                   double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out value);
        }

        private static string Format(double value)
        {
            return value.ToString("G10", CultureInfo.CurrentCulture);
        }

        private static string ToDisplayName(DropPanelRotationMode mode)
        {
            switch (mode)
            {
                case DropPanelRotationMode.FollowColumnLocalAxis: return "Follow Column Local Axis";
                case DropPanelRotationMode.UserDefinedAngle: return "User-Defined Angle";
                default: return "Global X-Y";
            }
        }

        private static DropPanelRotationMode FromDisplayName(string value)
        {
            if (string.Equals(value, "Follow Column Local Axis", StringComparison.Ordinal))
            {
                return DropPanelRotationMode.FollowColumnLocalAxis;
            }

            return string.Equals(value, "User-Defined Angle", StringComparison.Ordinal)
                ? DropPanelRotationMode.UserDefinedAngle
                : DropPanelRotationMode.GlobalXY;
        }

        private static void ShowWarning(string message)
        {
            MessageBox.Show(message, "Drop Panel", MessageBoxButton.OK, MessageBoxImage.Warning);
        }

        private static void ShowError(string message)
        {
            MessageBox.Show(message, "Drop Panel", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }
}
