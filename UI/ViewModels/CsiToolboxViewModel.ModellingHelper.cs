using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Input;
using ExcelCSIToolBox.Core.Abstractions.CSI;
using ExcelCSIToolBox.Core.Common.Commands;
using ExcelCSIToolBox.Core.Common.Results;
using ExcelCSIToolBox.Data.CSISapModel.FrameObject;
using ExcelCSIToolBox.Data.CSISapModel.PointObject;
using ExcelCSIToolBoxAddIn.AddIn.Modules.ModellingHelpers;
using ExcelCSIToolBoxAddIn.UI.Views;

namespace ExcelCSIToolBoxAddIn.UI.ViewModels
{
    public partial class CsiToolboxViewModel
    {
        private string _point1Name;
        private string _point1X;
        private string _point1Y;
        private string _point1Z;
        private string _point2Name;
        private string _point2X;
        private string _point2Y;
        private string _point2Z;
        private string _referenceFrameName;
        private string _referenceFrameLength;
        private string _line1Name;
        private string _line2Name;
        private string _selectedSection;
        private bool _isAutoTrimExtendEnabled;
        private bool _includeStartEndConnectors;
        private string _selectedAdjustmentMode;
        private bool _isFixFixSelected;
        private bool _isPinPinSelected;
        private int _numberOfSpaces;
        private ModellingHelperActionRouter _modellingHelperActionRouter;

        private const string AdjustJEndToFirstIntersection = "Adjust J-End To First Intersection";
        private const string AdjustBothEndsToNearestIntersections = "Adjust Both Ends To Nearest Intersections";
        private const int ShowWindowRestore = 9;

        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        private void InitializeModellingHelperPage()
        {
            AvailableSections = new ObservableCollection<string>();
            AdjustmentModes = new ObservableCollection<string>
            {
                AdjustJEndToFirstIntersection,
                AdjustBothEndsToNearestIntersections
            };
            _selectedAdjustmentMode = AdjustJEndToFirstIntersection;
            _isFixFixSelected = true;
            _numberOfSpaces = 3;

            _modellingHelperActionRouter = new ModellingHelperActionRouter()
                .Register(ModellingHelperActionKeys.OpenCreateArrayPerpendicularToPath, OpenCreateArrayPerpendicularToPathWindow)
                .Register(ModellingHelperActionKeys.OpenArrayBetweenTwoLines, OpenArrayBetweenTwoLinesWindow)
                .Register(ModellingHelperActionKeys.PickPoint1, PickPoint1)
                .Register(ModellingHelperActionKeys.PickPoint2, PickPoint2)
                .Register(ModellingHelperActionKeys.PickReferenceFrame, PickReferenceFrame)
                .Register(ModellingHelperActionKeys.PickLine1, PickLine1)
                .Register(ModellingHelperActionKeys.PickLine2, PickLine2)
                .Register(ModellingHelperActionKeys.CreateFrames, CreateFrames)
                .Register(ModellingHelperActionKeys.CreateArrayBetweenTwoLinesFrames, CreateArrayBetweenTwoLinesFrames);

            OpenCreateArrayPerpendicularToPathWindowCommand = new RelayCommand(() => ExecuteModellingHelperAction(ModellingHelperActionKeys.OpenCreateArrayPerpendicularToPath), CanExecuteCsiAction);
            OpenArrayBetweenTwoLinesWindowCommand = new RelayCommand(() => ExecuteModellingHelperAction(ModellingHelperActionKeys.OpenArrayBetweenTwoLines), CanExecuteCsiAction);
            PickPoint1Command = new RelayCommand(() => ExecuteModellingHelperAction(ModellingHelperActionKeys.PickPoint1), CanExecuteCsiAction);
            PickPoint2Command = new RelayCommand(() => ExecuteModellingHelperAction(ModellingHelperActionKeys.PickPoint2), CanExecuteCsiAction);
            PickReferenceFrameCommand = new RelayCommand(() => ExecuteModellingHelperAction(ModellingHelperActionKeys.PickReferenceFrame), CanExecuteCsiAction);
            PickLine1Command = new RelayCommand(() => ExecuteModellingHelperAction(ModellingHelperActionKeys.PickLine1), CanExecuteCsiAction);
            PickLine2Command = new RelayCommand(() => ExecuteModellingHelperAction(ModellingHelperActionKeys.PickLine2), CanExecuteCsiAction);
            CreateFramesCommand = new RelayCommand(() => ExecuteModellingHelperAction(ModellingHelperActionKeys.CreateFrames), CanExecuteCsiAction);
            CreateArrayBetweenTwoLinesFramesCommand = new RelayCommand(() => ExecuteModellingHelperAction(ModellingHelperActionKeys.CreateArrayBetweenTwoLinesFrames), CanExecuteCsiAction);
            CloseWindowCommand = new RelayCommand<Window>(CloseWindow);
            InitializeOffsetFromSetOfLinesPage();
        }

        private void ExecuteModellingHelperAction(string key)
        {
            _modellingHelperActionRouter.Execute(key);
        }

        public string Point1Name
        {
            get { return _point1Name; }
            set
            {
                if (_point1Name == value)
                {
                    return;
                }

                _point1Name = value;
                OnPropertyChanged();
                ClearPoint1Coordinates();
            }
        }

        public string Point1X
        {
            get { return _point1X; }
            set
            {
                if (_point1X == value)
                {
                    return;
                }

                _point1X = value;
                OnPropertyChanged();
            }
        }

        public string Point1Y
        {
            get { return _point1Y; }
            set
            {
                if (_point1Y == value)
                {
                    return;
                }

                _point1Y = value;
                OnPropertyChanged();
            }
        }

        public string Point1Z
        {
            get { return _point1Z; }
            set
            {
                if (_point1Z == value)
                {
                    return;
                }

                _point1Z = value;
                OnPropertyChanged();
            }
        }

        public string Point2Name
        {
            get { return _point2Name; }
            set
            {
                if (_point2Name == value)
                {
                    return;
                }

                _point2Name = value;
                OnPropertyChanged();
                ClearPoint2Coordinates();
            }
        }

        public string Point2X
        {
            get { return _point2X; }
            set
            {
                if (_point2X == value)
                {
                    return;
                }

                _point2X = value;
                OnPropertyChanged();
            }
        }

        public string Point2Y
        {
            get { return _point2Y; }
            set
            {
                if (_point2Y == value)
                {
                    return;
                }

                _point2Y = value;
                OnPropertyChanged();
            }
        }

        public string Point2Z
        {
            get { return _point2Z; }
            set
            {
                if (_point2Z == value)
                {
                    return;
                }

                _point2Z = value;
                OnPropertyChanged();
            }
        }

        public string ReferenceFrameName
        {
            get { return _referenceFrameName; }
            set
            {
                if (_referenceFrameName == value)
                {
                    return;
                }

                _referenceFrameName = value;
                OnPropertyChanged();
                ReferenceFrameLength = string.Empty;
            }
        }

        public string ReferenceFrameLength
        {
            get { return _referenceFrameLength; }
            set
            {
                if (_referenceFrameLength == value)
                {
                    return;
                }

                _referenceFrameLength = value;
                OnPropertyChanged();
            }
        }

        public string Line1Name
        {
            get { return _line1Name; }
            set
            {
                if (_line1Name == value)
                {
                    return;
                }

                _line1Name = value;
                OnPropertyChanged();
            }
        }

        public string Line2Name
        {
            get { return _line2Name; }
            set
            {
                if (_line2Name == value)
                {
                    return;
                }

                _line2Name = value;
                OnPropertyChanged();
            }
        }

        public ObservableCollection<string> AvailableSections { get; private set; }

        public string SelectedSection
        {
            get { return _selectedSection; }
            set
            {
                if (_selectedSection == value)
                {
                    return;
                }

                _selectedSection = value;
                OnPropertyChanged();
            }
        }

        public bool IsAutoTrimExtendEnabled
        {
            get { return _isAutoTrimExtendEnabled; }
            set
            {
                if (_isAutoTrimExtendEnabled == value)
                {
                    return;
                }

                _isAutoTrimExtendEnabled = value;
                OnPropertyChanged();
            }
        }

        public bool IncludeStartEndConnectors
        {
            get { return _includeStartEndConnectors; }
            set
            {
                if (_includeStartEndConnectors == value)
                {
                    return;
                }

                _includeStartEndConnectors = value;
                OnPropertyChanged();
            }
        }

        public ObservableCollection<string> AdjustmentModes { get; private set; }

        public string SelectedAdjustmentMode
        {
            get { return _selectedAdjustmentMode; }
            set
            {
                if (_selectedAdjustmentMode == value)
                {
                    return;
                }

                _selectedAdjustmentMode = value;
                OnPropertyChanged();
            }
        }

        public bool IsFixFixSelected
        {
            get { return _isFixFixSelected; }
            set
            {
                if (_isFixFixSelected == value)
                {
                    return;
                }

                _isFixFixSelected = value;
                OnPropertyChanged();
                if (value && _isPinPinSelected)
                {
                    _isPinPinSelected = false;
                    OnPropertyChanged(nameof(IsPinPinSelected));
                }
            }
        }

        public bool IsPinPinSelected
        {
            get { return _isPinPinSelected; }
            set
            {
                if (_isPinPinSelected == value)
                {
                    return;
                }

                _isPinPinSelected = value;
                OnPropertyChanged();
                if (value && _isFixFixSelected)
                {
                    _isFixFixSelected = false;
                    OnPropertyChanged(nameof(IsFixFixSelected));
                }
            }
        }

        public int NumberOfSpaces
        {
            get { return _numberOfSpaces; }
            set
            {
                int normalizedValue = value < 1 ? 1 : value;
                if (_numberOfSpaces == normalizedValue)
                {
                    return;
                }

                _numberOfSpaces = normalizedValue;
                OnPropertyChanged();
            }
        }

        public ICommand OpenCreateArrayPerpendicularToPathWindowCommand { get; private set; }
        public ICommand OpenArrayBetweenTwoLinesWindowCommand { get; private set; }
        public ICommand PickPoint1Command { get; private set; }
        public ICommand PickPoint2Command { get; private set; }
        public ICommand PickReferenceFrameCommand { get; private set; }
        public ICommand PickLine1Command { get; private set; }
        public ICommand PickLine2Command { get; private set; }
        public ICommand CreateFramesCommand { get; private set; }
        public ICommand CreateArrayBetweenTwoLinesFramesCommand { get; private set; }
        public ICommand CloseWindowCommand { get; private set; }

        private void OpenCreateArrayPerpendicularToPathWindow()
        {
            RefreshAvailableFrameSectionsForArrayTool();

            var window = new CreateArrayPerpendicularToPathWindow(this);
            Window owner = GetActiveOwnerWindow();
            if (owner != null && !ReferenceEquals(owner, window))
            {
                window.Owner = owner;
                window.WindowStartupLocation = WindowStartupLocation.CenterOwner;
            }

            window.ShowDialog();
        }

        private void OpenArrayBetweenTwoLinesWindow()
        {
            RefreshAvailableFrameSectionsForArrayTool();

            var window = new ArrayBetweenTwoLinesWindow(this);
            Window owner = GetActiveOwnerWindow();
            if (owner != null && !ReferenceEquals(owner, window))
            {
                window.Owner = owner;
                window.WindowStartupLocation = WindowStartupLocation.CenterOwner;
            }

            window.ShowDialog();
        }

        private static Window GetActiveOwnerWindow()
        {
            if (Application.Current == null)
            {
                return null;
            }

            foreach (Window window in Application.Current.Windows)
            {
                if (window != null && window.IsActive)
                {
                    return window;
                }
            }

            return Application.Current.MainWindow;
        }

        private void RefreshAvailableFrameSectionsForArrayTool()
        {
            AvailableSections.Clear();

            if (!IsConnected)
            {
                return;
            }

            OperationResult<System.Collections.Generic.IReadOnlyList<string>> result = _csiConnectionService.GetFrameSectionNames();
            if (!result.IsSuccess || result.Data == null)
            {
                ShowWarning(result.Message);
                return;
            }

            foreach (string sectionName in result.Data)
            {
                if (!string.IsNullOrWhiteSpace(sectionName))
                {
                    AvailableSections.Add(sectionName);
                }
            }

            if (string.IsNullOrWhiteSpace(SelectedSection) && AvailableSections.Count > 0)
            {
                SelectedSection = AvailableSections[0];
            }
        }

        private void PickPoint1()
        {
            PickPoint(
                "Select one point object in ETABS.",
                delegate (PointObjectInfo point)
                {
                    Point1Name = point.Name;
                    Point1X = FormatCoordinate(point.X);
                    Point1Y = FormatCoordinate(point.Y);
                    Point1Z = FormatCoordinate(point.Z);
                });
        }

        private void PickPoint2()
        {
            PickPoint(
                "Select one point object in ETABS.",
                delegate (PointObjectInfo point)
                {
                    Point2Name = point.Name;
                    Point2X = FormatCoordinate(point.X);
                    Point2Y = FormatCoordinate(point.Y);
                    Point2Z = FormatCoordinate(point.Z);
                });
        }

        private void PickReferenceFrame()
        {
            CsiSelectedObjectDto selectedObject;
            if (!TryPickObjectInteractively(
                "Frame",
                "Pick Reference Frame",
                "Select one frame object in ETABS.",
                "Waiting for one frame object...",
                "Only one frame object should be selected.",
                "Selected object must be a frame object.",
                out selectedObject))
            {
                return;
            }

            ReferenceFrameName = selectedObject.UniqueName;
            OperationResult<double> lengthResult = GetFrameLength(selectedObject.UniqueName);
            ReferenceFrameLength = lengthResult.IsSuccess
                ? FormatCoordinate(lengthResult.Data)
                : string.Empty;
        }

        private void PickPoint(string noSelectionMessage, System.Action<PointObjectInfo> applyPoint)
        {
            CsiSelectedObjectDto selectedObject;
            if (!TryPickObjectInteractively(
                "Point",
                "Pick Point",
                noSelectionMessage,
                "Waiting for one point object...",
                "Only one point object should be selected.",
                "Selected object must be a point object.",
                out selectedObject))
            {
                return;
            }

            OperationResult<PointObjectInfo> coordinateResult = _csiConnectionService.GetPointCoordinates(selectedObject.UniqueName);
            if (!coordinateResult.IsSuccess || coordinateResult.Data == null)
            {
                ShowWarning(coordinateResult.Message);
                return;
            }

            applyPoint(coordinateResult.Data);
        }

        private void PickLine1()
        {
            CsiSelectedObjectDto selectedObject;
            if (TryPickObjectInteractively(
                "Frame",
                "Pick Line 1",
                "Select one frame object in ETABS.",
                "Waiting for one frame object...",
                "Only one frame object should be selected.",
                "Selected object must be a frame object.",
                out selectedObject))
            {
                Line1Name = selectedObject.UniqueName;
            }
        }

        private void PickLine2()
        {
            CsiSelectedObjectDto selectedObject;
            if (TryPickObjectInteractively(
                "Frame",
                "Pick Line 2",
                "Select one frame object in ETABS.",
                "Waiting for one frame object...",
                "Only one frame object should be selected.",
                "Selected object must be a frame object.",
                out selectedObject))
            {
                Line2Name = selectedObject.UniqueName;
            }
        }

        private bool TryGetCurrentSingleSelectedFrame(out CsiSelectedObjectDto selectedObject)
        {
            selectedObject = null;
            OperationResult<IReadOnlyList<CsiSelectedObjectDto>> selectedResult =
                _csiConnectionService.GetSelectedObjectsFromActiveModel();

            if (!selectedResult.IsSuccess)
            {
                string message = selectedResult.Message ?? string.Empty;
                ShowWarning(message.IndexOf("No running", StringComparison.OrdinalIgnoreCase) >= 0
                    ? selectedResult.Message
                    : "Please select one frame object in ETABS before clicking Pick.");
                return false;
            }

            if (selectedResult.Data == null || selectedResult.Data.Count == 0)
            {
                ShowWarning("Please select one frame object in ETABS before clicking Pick.");
                return false;
            }

            if (selectedResult.Data.Count > 1)
            {
                ShowWarning("Only one frame object should be selected.");
                return false;
            }

            CsiSelectedObjectDto selected = selectedResult.Data[0];
            if (selected == null || !string.Equals(selected.ObjectType, "Frame", StringComparison.OrdinalIgnoreCase))
            {
                ShowWarning("Selected object must be a frame object.");
                return false;
            }

            selectedObject = selected;
            return true;
        }

        private bool TryPickObjectInteractively(
            string requiredObjectType,
            string title,
            string instruction,
            string waitingMessage,
            string multipleSelectionMessage,
            string wrongTypeMessage,
            out CsiSelectedObjectDto selectedObject)
        {
            selectedObject = null;

            OperationResult clearResult = _csiConnectionService.ClearSelection();
            if (!clearResult.IsSuccess)
            {
                ShowWarning(clearResult.Message);
                return false;
            }

            var window = new InteractiveSelectionWindow(
                title,
                instruction,
                waitingMessage,
                multipleSelectionMessage,
                wrongTypeMessage,
                requiredObjectType,
                () => _csiConnectionService.GetSelectedObjectsFromActiveModel());

            Window owner = GetActiveOwnerWindow();
            if (owner != null && !ReferenceEquals(owner, window))
            {
                window.Owner = owner;
                window.WindowStartupLocation = WindowStartupLocation.CenterOwner;
            }

            ActivateConnectedCsiWindow();
            bool? result = window.ShowDialog();
            if (result != true || window.SelectedObject == null)
            {
                return false;
            }

            selectedObject = window.SelectedObject;
            return true;
        }

        private void ActivateConnectedCsiWindow()
        {
            try
            {
                OperationResult<ExcelCSIToolBox.Data.DTOs.CSI.CSISapModelConnectionInfoDTO> connectionResult =
                    _csiConnectionService.GetCurrentConnection();
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
                // Window activation is a convenience only; selection polling still works without it.
            }
        }

        private bool TryGetSingleSelectedObject(
            OperationResult<System.Collections.Generic.IReadOnlyList<CsiSelectedObjectDto>> selectedResult,
            string requiredObjectType,
            string noSelectionMessage,
            string multipleSelectionMessage,
            string wrongTypeMessage,
            out CsiSelectedObjectDto selectedObject)
        {
            selectedObject = null;

            if (!selectedResult.IsSuccess)
            {
                string message = selectedResult.Message ?? string.Empty;
                ShowWarning(message.IndexOf("No running", System.StringComparison.OrdinalIgnoreCase) >= 0
                    ? selectedResult.Message
                    : noSelectionMessage);
                return false;
            }

            if (selectedResult.Data == null || selectedResult.Data.Count == 0)
            {
                ShowWarning(noSelectionMessage);
                return false;
            }

            int matchingObjectCount = 0;
            CsiSelectedObjectDto matchingObject = null;
            foreach (CsiSelectedObjectDto item in selectedResult.Data)
            {
                if (item == null)
                {
                    continue;
                }

                if (string.Equals(item.ObjectType, requiredObjectType, System.StringComparison.OrdinalIgnoreCase))
                {
                    matchingObjectCount++;
                    matchingObject = item;
                }
            }

            if (matchingObjectCount > 1)
            {
                ShowWarning(multipleSelectionMessage);
                return false;
            }

            if (matchingObjectCount == 0)
            {
                ShowWarning(selectedResult.Data.Count == 0 ? noSelectionMessage : wrongTypeMessage);
                return false;
            }

            selectedObject = matchingObject;
            return true;
        }

        private void CreateFrames()
        {
            OperationResult<ArrayFrameInput> inputResult = ReadArrayFrameInput();
            if (!inputResult.IsSuccess)
            {
                ShowWarning(inputResult.Message);
                return;
            }

            ArrayFrameInput input = inputResult.Data;
            Vector3 generatedDirection = Subtract(input.Point2, input.Point1);
            Vector3 arrayDirection = Subtract(input.ReferencePointJ, input.ReferencePointI);
            double generatedLength = Length(generatedDirection);
            double referenceLength = Length(arrayDirection);
            if (generatedLength <= 0 || referenceLength <= 0)
            {
                ShowWarning("Point 1, Point 2, and Reference Frame must define non-zero lengths.");
                return;
            }

            double dot = Dot(Normalize(generatedDirection), Normalize(arrayDirection));
            if (System.Math.Abs(dot) > 0.001)
            {
                ShowWarning("The line from Point 1 to Point 2 must be perpendicular to the selected Reference Frame.");
                return;
            }

            WorkingPlane workingPlane = DetectWorkingPlane(generatedDirection, arrayDirection);
            double searchRadius = GetDefaultSearchRadius();
            double tolerance = GetGeometryTolerance(searchRadius);
            List<ModelSegment> modelSegments = IsAutoTrimExtendEnabled
                ? LoadExistingTargetSegments(ReferenceFrameName)
                : new List<ModelSegment>();

            Vector3 step = Scale(Normalize(arrayDirection), referenceLength / NumberOfSpaces);
            var frameRequests = new List<FrameAddRequestDto>();
            var trimExtendSummary = new TrimExtendSummary();
            for (int i = 0; i <= NumberOfSpaces; i++)
            {
                Vector3 offset = Scale(step, i);
                Vector3 start = Add(input.Point1, offset);
                Vector3 end = Add(input.Point2, offset);

                if (IsAutoTrimExtendEnabled)
                {
                    CandidateAdjustmentResult adjustment = ApplyAutoTrimExtend(
                        start,
                        end,
                        modelSegments,
                        workingPlane,
                        searchRadius,
                        tolerance);
                    trimExtendSummary.Add(adjustment);

                    if (adjustment.IsSkipped)
                    {
                        continue;
                    }

                    start = adjustment.Start;
                    end = adjustment.End;
                }

                ReorderFrameEnds(ref start, ref end);

                frameRequests.Add(new FrameAddRequestDto
                {
                    SectionName = SelectedSection,
                    Xi = start.X,
                    Yi = start.Y,
                    Zi = start.Z,
                    Xj = end.X,
                    Yj = end.Y,
                    Zj = end.Z
                });
            }

            if (frameRequests.Count == 0)
            {
                return;
            }

            OperationResult<FrameAddBatchResultDto> addResult = _csiConnectionService.AddFrameObjects(
                new FrameAddBatchRequestDto
                {
                    Frames = frameRequests,
                    SuppressViewRefresh = true
                });
            if (!addResult.IsSuccess || addResult.Data == null)
            {
                return;
            }

            FrameAddBatchResultDto batch = addResult.Data;
            if (batch.SuccessCount == 0)
            {
                return;
            }

            if (IsPinPinSelected)
            {
                bool[] momentReleases = new[] { false, false, false, false, true, true };
                OperationResult releaseResult = _csiConnectionService.SetFrameReleases(
                    batch.SuccessfulFrameNames,
                    momentReleases,
                    momentReleases,
                    suppressViewRefresh: true);
                if (!releaseResult.IsSuccess)
                {
                    return;
                }
            }

            _csiConnectionService.RefreshView(false);
        }

        private void CreateArrayBetweenTwoLinesFrames()
        {
            OperationResult<ArrayBetweenTwoLinesInput> inputResult = ReadArrayBetweenTwoLinesInput();
            if (!inputResult.IsSuccess)
            {
                ShowWarning(inputResult.Message);
                return;
            }

            ArrayBetweenTwoLinesInput input = inputResult.Data;
            Vector3 line1Direction = Subtract(input.Line1End, input.Line1Start);
            Vector3 line2Direction = Subtract(input.Line2End, input.Line2Start);
            double line1Length = Length(line1Direction);
            double line2Length = Length(line2Direction);
            double tolerance = Math.Max(0.000001, Math.Max(line1Length, line2Length) * 0.000001);

            if (line1Length <= tolerance || line2Length <= tolerance)
            {
                ShowWarning("Both selected lines must have valid length.");
                return;
            }

            Vector3 line2Start = input.Line2Start;
            Vector3 line2End = input.Line2End;
            double distanceOption1 = Distance(input.Line1Start, line2Start) + Distance(input.Line1End, line2End);
            double distanceOption2 = Distance(input.Line1Start, line2End) + Distance(input.Line1End, line2Start);
            if (distanceOption2 < distanceOption1)
            {
                Vector3 temp = line2Start;
                line2Start = line2End;
                line2End = temp;
            }

            int startIndex = IncludeStartEndConnectors ? 0 : 1;
            int endIndex = IncludeStartEndConnectors ? NumberOfSpaces : NumberOfSpaces - 1;
            var frameRequests = new List<FrameAddRequestDto>();
            int skippedCount = 0;

            for (int i = startIndex; i <= endIndex; i++)
            {
                double t = (double)i / NumberOfSpaces;
                Vector3 start = Interpolate(input.Line1Start, input.Line1End, t);
                Vector3 end = Interpolate(line2Start, line2End, t);
                if (Distance(start, end) <= tolerance)
                {
                    skippedCount++;
                    continue;
                }

                ReorderFrameEnds(ref start, ref end);

                frameRequests.Add(new FrameAddRequestDto
                {
                    SectionName = SelectedSection,
                    Xi = start.X,
                    Yi = start.Y,
                    Zi = start.Z,
                    Xj = end.X,
                    Yj = end.Y,
                    Zj = end.Z
                });
            }

            OperationResult<FrameAddBatchResultDto> addResult = _csiConnectionService.AddFrameObjects(
                new FrameAddBatchRequestDto
                {
                    Frames = frameRequests,
                    SuppressViewRefresh = true
                });
            if (!addResult.IsSuccess || addResult.Data == null)
            {
                return;
            }

            FrameAddBatchResultDto batch = addResult.Data;
            if (IsPinPinSelected && batch.SuccessCount > 0)
            {
                bool[] momentReleases = new[] { false, false, false, false, true, true };
                _csiConnectionService.SetFrameReleases(
                    batch.SuccessfulFrameNames,
                    momentReleases,
                    momentReleases,
                    suppressViewRefresh: true);
            }

            _csiConnectionService.RefreshView(false);
            skippedCount += batch.FailureCount;
            MessageBox.Show(
                "Created frames: " + batch.SuccessCount.ToString(CultureInfo.CurrentCulture) + "\n" +
                "Skipped frames: " + skippedCount.ToString(CultureInfo.CurrentCulture),
                ProductTitle,
                MessageBoxButton.OK,
                MessageBoxImage.Information);
        }

        private OperationResult<ArrayBetweenTwoLinesInput> ReadArrayBetweenTwoLinesInput()
        {
            if (!IsConnected)
            {
                return OperationResult<ArrayBetweenTwoLinesInput>.Failure("ETABS model is not connected.");
            }

            if (string.IsNullOrWhiteSpace(Line1Name))
            {
                return OperationResult<ArrayBetweenTwoLinesInput>.Failure("Line 1 is required.");
            }

            if (string.IsNullOrWhiteSpace(Line2Name))
            {
                return OperationResult<ArrayBetweenTwoLinesInput>.Failure("Line 2 is required.");
            }

            if (string.Equals(Line1Name, Line2Name, StringComparison.OrdinalIgnoreCase))
            {
                return OperationResult<ArrayBetweenTwoLinesInput>.Failure("Line 1 and Line 2 must be different frames.");
            }

            if (string.IsNullOrWhiteSpace(SelectedSection))
            {
                return OperationResult<ArrayBetweenTwoLinesInput>.Failure("Assign Section is required.");
            }

            if (NumberOfSpaces < 1)
            {
                return OperationResult<ArrayBetweenTwoLinesInput>.Failure("Number of Spaces must be at least 1.");
            }

            if (!IncludeStartEndConnectors && NumberOfSpaces == 1)
            {
                return OperationResult<ArrayBetweenTwoLinesInput>.Failure("Number of Spaces must be greater than 1 when Start / End Connectors are not included.");
            }

            OperationResult<FrameEndPointInfo> line1PointsResult = _csiConnectionService.GetFramePoints(Line1Name);
            if (!line1PointsResult.IsSuccess || line1PointsResult.Data == null)
            {
                return OperationResult<ArrayBetweenTwoLinesInput>.Failure("Cannot read Line 1 endpoints: " + line1PointsResult.Message);
            }

            OperationResult<FrameEndPointInfo> line2PointsResult = _csiConnectionService.GetFramePoints(Line2Name);
            if (!line2PointsResult.IsSuccess || line2PointsResult.Data == null)
            {
                return OperationResult<ArrayBetweenTwoLinesInput>.Failure("Cannot read Line 2 endpoints: " + line2PointsResult.Message);
            }

            OperationResult<PointObjectInfo> line1StartResult = _csiConnectionService.GetPointCoordinates(line1PointsResult.Data.PointI);
            OperationResult<PointObjectInfo> line1EndResult = _csiConnectionService.GetPointCoordinates(line1PointsResult.Data.PointJ);
            OperationResult<PointObjectInfo> line2StartResult = _csiConnectionService.GetPointCoordinates(line2PointsResult.Data.PointI);
            OperationResult<PointObjectInfo> line2EndResult = _csiConnectionService.GetPointCoordinates(line2PointsResult.Data.PointJ);

            if (!line1StartResult.IsSuccess || line1StartResult.Data == null ||
                !line1EndResult.IsSuccess || line1EndResult.Data == null ||
                !line2StartResult.IsSuccess || line2StartResult.Data == null ||
                !line2EndResult.IsSuccess || line2EndResult.Data == null)
            {
                return OperationResult<ArrayBetweenTwoLinesInput>.Failure("Cannot read selected line endpoint coordinates.");
            }

            return OperationResult<ArrayBetweenTwoLinesInput>.Success(new ArrayBetweenTwoLinesInput
            {
                Line1Start = ToVector(line1StartResult.Data),
                Line1End = ToVector(line1EndResult.Data),
                Line2Start = ToVector(line2StartResult.Data),
                Line2End = ToVector(line2EndResult.Data)
            });
        }

        private OperationResult<ArrayFrameInput> ReadArrayFrameInput()
        {
            if (string.IsNullOrWhiteSpace(Point1Name))
            {
                return OperationResult<ArrayFrameInput>.Failure("Point 1 is required.");
            }

            if (string.IsNullOrWhiteSpace(Point2Name))
            {
                return OperationResult<ArrayFrameInput>.Failure("Point 2 is required.");
            }

            if (string.IsNullOrWhiteSpace(ReferenceFrameName))
            {
                return OperationResult<ArrayFrameInput>.Failure("Reference Frame is required.");
            }

            if (string.IsNullOrWhiteSpace(SelectedSection))
            {
                return OperationResult<ArrayFrameInput>.Failure("Assign Section is required.");
            }

            OperationResult<PointObjectInfo> point1Result = _csiConnectionService.GetPointCoordinates(Point1Name);
            if (!point1Result.IsSuccess || point1Result.Data == null)
            {
                return OperationResult<ArrayFrameInput>.Failure("Cannot read Point 1 coordinates: " + point1Result.Message);
            }

            OperationResult<PointObjectInfo> point2Result = _csiConnectionService.GetPointCoordinates(Point2Name);
            if (!point2Result.IsSuccess || point2Result.Data == null)
            {
                return OperationResult<ArrayFrameInput>.Failure("Cannot read Point 2 coordinates: " + point2Result.Message);
            }

            OperationResult<FrameEndPointInfo> referencePointsResult = _csiConnectionService.GetFramePoints(ReferenceFrameName);
            if (!referencePointsResult.IsSuccess || referencePointsResult.Data == null)
            {
                return OperationResult<ArrayFrameInput>.Failure("Cannot read Reference Frame endpoints: " + referencePointsResult.Message);
            }

            OperationResult<PointObjectInfo> referencePointIResult = _csiConnectionService.GetPointCoordinates(referencePointsResult.Data.PointI);
            if (!referencePointIResult.IsSuccess || referencePointIResult.Data == null)
            {
                return OperationResult<ArrayFrameInput>.Failure("Cannot read Reference Frame I-End coordinates: " + referencePointIResult.Message);
            }

            OperationResult<PointObjectInfo> referencePointJResult = _csiConnectionService.GetPointCoordinates(referencePointsResult.Data.PointJ);
            if (!referencePointJResult.IsSuccess || referencePointJResult.Data == null)
            {
                return OperationResult<ArrayFrameInput>.Failure("Cannot read Reference Frame J-End coordinates: " + referencePointJResult.Message);
            }

            Point1X = FormatCoordinate(point1Result.Data.X);
            Point1Y = FormatCoordinate(point1Result.Data.Y);
            Point1Z = FormatCoordinate(point1Result.Data.Z);
            Point2X = FormatCoordinate(point2Result.Data.X);
            Point2Y = FormatCoordinate(point2Result.Data.Y);
            Point2Z = FormatCoordinate(point2Result.Data.Z);
            ReferenceFrameLength = FormatCoordinate(Length(Subtract(ToVector(referencePointJResult.Data), ToVector(referencePointIResult.Data))));

            return OperationResult<ArrayFrameInput>.Success(new ArrayFrameInput
            {
                Point1 = ToVector(point1Result.Data),
                Point2 = ToVector(point2Result.Data),
                ReferencePointI = ToVector(referencePointIResult.Data),
                ReferencePointJ = ToVector(referencePointJResult.Data)
            });
        }

        private OperationResult<double> GetFrameLength(string frameName)
        {
            OperationResult<FrameEndPointInfo> pointsResult = _csiConnectionService.GetFramePoints(frameName);
            if (!pointsResult.IsSuccess || pointsResult.Data == null)
            {
                return OperationResult<double>.Failure(pointsResult.Message);
            }

            OperationResult<PointObjectInfo> pointIResult = _csiConnectionService.GetPointCoordinates(pointsResult.Data.PointI);
            OperationResult<PointObjectInfo> pointJResult = _csiConnectionService.GetPointCoordinates(pointsResult.Data.PointJ);
            if (!pointIResult.IsSuccess || pointIResult.Data == null)
            {
                return OperationResult<double>.Failure(pointIResult.Message);
            }

            if (!pointJResult.IsSuccess || pointJResult.Data == null)
            {
                return OperationResult<double>.Failure(pointJResult.Message);
            }

            return OperationResult<double>.Success(Length(Subtract(ToVector(pointJResult.Data), ToVector(pointIResult.Data))));
        }

        private static void ReorderFrameEnds(ref Vector3 start, ref Vector3 end)
        {
            Vector3 delta = Subtract(end, start);
            double horizontal = System.Math.Max(System.Math.Abs(delta.X), System.Math.Abs(delta.Y));
            bool mainlyVertical = System.Math.Abs(delta.Z) > horizontal;
            bool shouldSwap = mainlyVertical
                ? start.Z > end.Z
                : start.X > end.X || (NearlyEqual(start.X, end.X) && start.Y > end.Y);

            if (shouldSwap)
            {
                Vector3 temp = start;
                start = end;
                end = temp;
            }
        }

        private static bool NearlyEqual(double left, double right)
        {
            return System.Math.Abs(left - right) < 0.000001;
        }

        private static Vector3 ToVector(PointObjectInfo point)
        {
            return new Vector3(point.X, point.Y, point.Z);
        }

        private static Vector3 Add(Vector3 left, Vector3 right)
        {
            return new Vector3(left.X + right.X, left.Y + right.Y, left.Z + right.Z);
        }

        private static Vector3 Interpolate(Vector3 start, Vector3 end, double t)
        {
            return Add(start, Scale(Subtract(end, start), t));
        }

        private static Vector3 Subtract(Vector3 left, Vector3 right)
        {
            return new Vector3(left.X - right.X, left.Y - right.Y, left.Z - right.Z);
        }

        private static Vector3 Scale(Vector3 value, double scale)
        {
            return new Vector3(value.X * scale, value.Y * scale, value.Z * scale);
        }

        private static double Dot(Vector3 left, Vector3 right)
        {
            return left.X * right.X + left.Y * right.Y + left.Z * right.Z;
        }

        private static double Length(Vector3 value)
        {
            return System.Math.Sqrt(Dot(value, value));
        }

        private static Vector3 Normalize(Vector3 value)
        {
            double length = Length(value);
            return length <= 0 ? new Vector3(0, 0, 0) : Scale(value, 1 / length);
        }

        private CandidateAdjustmentResult ApplyAutoTrimExtend(
            Vector3 originalStart,
            Vector3 originalEnd,
            IReadOnlyList<ModelSegment> modelSegments,
            WorkingPlane workingPlane,
            double searchRadius,
            double tolerance)
        {
            double originalLength = Distance(originalStart, originalEnd);
            if (originalLength <= tolerance)
            {
                return CandidateAdjustmentResult.Skipped(originalStart, originalEnd);
            }

            var intersections = FindValidIntersections(
                originalStart,
                originalEnd,
                modelSegments,
                workingPlane,
                searchRadius,
                tolerance);

            Vector3 adjustedStart = originalStart;
            Vector3 adjustedEnd = originalEnd;

            if (string.Equals(SelectedAdjustmentMode, AdjustBothEndsToNearestIntersections, StringComparison.OrdinalIgnoreCase))
            {
                IntersectionResult targetI = GetNearestIntersectionForIEnd(intersections, originalLength, tolerance);
                IntersectionResult targetJ = GetNearestIntersectionForJEnd(intersections, originalLength, tolerance);

                if (targetI != null)
                {
                    adjustedStart = targetI.Point;
                }

                if (targetJ != null)
                {
                    adjustedEnd = targetJ.Point;
                }
            }
            else
            {
                IntersectionResult target = GetNearestIntersectionForJEnd(intersections, originalLength, tolerance);
                if (target != null)
                {
                    adjustedEnd = target.Point;
                }
            }

            double adjustedLength = Distance(adjustedStart, adjustedEnd);
            if (adjustedLength <= tolerance)
            {
                return CandidateAdjustmentResult.Skipped(adjustedStart, adjustedEnd);
            }

            double signedDirection = Dot(Subtract(adjustedEnd, adjustedStart), Subtract(originalEnd, originalStart));
            if (signedDirection <= tolerance)
            {
                return CandidateAdjustmentResult.Skipped(adjustedStart, adjustedEnd);
            }

            return new CandidateAdjustmentResult
            {
                Start = adjustedStart,
                End = adjustedEnd,
                IsSkipped = false,
                IsTrimmed = adjustedLength < originalLength - tolerance,
                IsExtended = adjustedLength > originalLength + tolerance
            };
        }

        private List<IntersectionResult> FindValidIntersections(
            Vector3 candidateStart,
            Vector3 candidateEnd,
            IReadOnlyList<ModelSegment> modelSegments,
            WorkingPlane workingPlane,
            double searchRadius,
            double tolerance)
        {
            var intersections = new List<IntersectionResult>();
            BoundingBox corridor = BoundingBox.FromSegment(candidateStart, candidateEnd, searchRadius);
            Point2D candidateA = Project(candidateStart, workingPlane);
            Point2D candidateB = Project(candidateEnd, workingPlane);
            Vector2D candidateDirection = Subtract(candidateB, candidateA);
            double candidateDirectionLengthSquared = Dot(candidateDirection, candidateDirection);
            if (candidateDirectionLengthSquared <= tolerance * tolerance)
            {
                return intersections;
            }

            double constantCoordinate = GetPlaneConstant(candidateStart, workingPlane);
            foreach (ModelSegment segment in modelSegments)
            {
                if (segment == null || !corridor.Intersects(segment.Bounds))
                {
                    continue;
                }

                if (!IsInSameWorkingPlane(segment.Start, segment.End, workingPlane, constantCoordinate, tolerance))
                {
                    continue;
                }

                Point2D targetA = Project(segment.Start, workingPlane);
                Point2D targetB = Project(segment.End, workingPlane);
                Intersection2DResult intersection2D;
                if (!TryIntersectLineWithSegment(candidateA, candidateDirection, targetA, targetB, tolerance, out intersection2D))
                {
                    continue;
                }

                if (intersection2D.CandidateParameter * Distance(candidateStart, candidateEnd) <= tolerance)
                {
                    continue;
                }

                if (Math.Abs(intersection2D.CandidateParameter - 1.0) * Distance(candidateStart, candidateEnd) <= tolerance)
                {
                    continue;
                }

                intersections.Add(new IntersectionResult
                {
                    Point = Unproject(intersection2D.Point, workingPlane, constantCoordinate),
                    CandidateParameter = intersection2D.CandidateParameter
                });
            }

            return intersections;
        }

        private static bool TryIntersectLineWithSegment(
            Point2D candidateOrigin,
            Vector2D candidateDirection,
            Point2D targetA,
            Point2D targetB,
            double tolerance,
            out Intersection2DResult result)
        {
            result = null;
            Vector2D targetDirection = Subtract(targetB, targetA);
            double denominator = Cross(candidateDirection, targetDirection);
            if (Math.Abs(denominator) <= tolerance)
            {
                return false;
            }

            Vector2D delta = Subtract(targetA, candidateOrigin);
            double candidateParameter = Cross(delta, targetDirection) / denominator;
            double targetParameter = Cross(delta, candidateDirection) / denominator;
            if (targetParameter < -tolerance || targetParameter > 1.0 + tolerance)
            {
                return false;
            }

            result = new Intersection2DResult
            {
                Point = new Point2D(
                    candidateOrigin.U + candidateParameter * candidateDirection.U,
                    candidateOrigin.V + candidateParameter * candidateDirection.V),
                CandidateParameter = candidateParameter
            };
            return true;
        }

        private static IntersectionResult GetNearestIntersectionForIEnd(
            IReadOnlyList<IntersectionResult> intersections,
            double originalLength,
            double tolerance)
        {
            IntersectionResult best = null;
            double bestDistance = double.MaxValue;
            foreach (IntersectionResult intersection in intersections)
            {
                if (intersection.CandidateParameter >= 1.0)
                {
                    continue;
                }

                double distanceFromI = Math.Abs(intersection.CandidateParameter) * originalLength;
                if (distanceFromI <= tolerance || Math.Abs(intersection.CandidateParameter - 1.0) * originalLength <= tolerance)
                {
                    continue;
                }

                if (distanceFromI < bestDistance)
                {
                    bestDistance = distanceFromI;
                    best = intersection;
                }
            }

            return best;
        }

        private static IntersectionResult GetNearestIntersectionForJEnd(
            IReadOnlyList<IntersectionResult> intersections,
            double originalLength,
            double tolerance)
        {
            IntersectionResult best = null;
            double bestDistance = double.MaxValue;
            foreach (IntersectionResult intersection in intersections)
            {
                if (intersection.CandidateParameter <= 0)
                {
                    continue;
                }

                double distanceFromJ = Math.Abs(intersection.CandidateParameter - 1.0) * originalLength;
                if (distanceFromJ <= tolerance || intersection.CandidateParameter * originalLength <= tolerance)
                {
                    continue;
                }

                if (distanceFromJ < bestDistance)
                {
                    bestDistance = distanceFromJ;
                    best = intersection;
                }
            }

            return best;
        }

        private List<ModelSegment> LoadExistingTargetSegments(string referenceFrameName)
        {
            var segments = new List<ModelSegment>();
            OperationResult<IReadOnlyList<string>> frameNamesResult = _csiConnectionService.GetFrameNames();
            if (frameNamesResult.IsSuccess && frameNamesResult.Data != null)
            {
                foreach (string frameName in frameNamesResult.Data)
                {
                    if (string.IsNullOrWhiteSpace(frameName) ||
                        string.Equals(frameName, referenceFrameName, StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    OperationResult<FrameEndPointInfo> framePointsResult = _csiConnectionService.GetFramePoints(frameName);
                    if (!framePointsResult.IsSuccess || framePointsResult.Data == null)
                    {
                        continue;
                    }

                    OperationResult<PointObjectInfo> pointIResult = _csiConnectionService.GetPointCoordinates(framePointsResult.Data.PointI);
                    OperationResult<PointObjectInfo> pointJResult = _csiConnectionService.GetPointCoordinates(framePointsResult.Data.PointJ);
                    if (!pointIResult.IsSuccess || pointIResult.Data == null ||
                        !pointJResult.IsSuccess || pointJResult.Data == null)
                    {
                        continue;
                    }

                    segments.Add(new ModelSegment(ToVector(pointIResult.Data), ToVector(pointJResult.Data), frameName));
                }
            }

            OperationResult<IReadOnlyList<string>> shellNamesResult = _csiConnectionService.GetShellNames();
            if (shellNamesResult.IsSuccess && shellNamesResult.Data != null)
            {
                foreach (string areaName in shellNamesResult.Data)
                {
                    if (string.IsNullOrWhiteSpace(areaName))
                    {
                        continue;
                    }

                    OperationResult<IReadOnlyList<string>> shellPointsResult = _csiConnectionService.GetShellPoints(areaName);
                    if (!shellPointsResult.IsSuccess || shellPointsResult.Data == null || shellPointsResult.Data.Count < 2)
                    {
                        continue;
                    }

                    var polygonPoints = new List<Vector3>();
                    foreach (string pointName in shellPointsResult.Data)
                    {
                        OperationResult<PointObjectInfo> pointResult = _csiConnectionService.GetPointCoordinates(pointName);
                        if (pointResult.IsSuccess && pointResult.Data != null)
                        {
                            polygonPoints.Add(ToVector(pointResult.Data));
                        }
                    }

                    for (int i = 0; i < polygonPoints.Count; i++)
                    {
                        Vector3 start = polygonPoints[i];
                        Vector3 end = polygonPoints[(i + 1) % polygonPoints.Count];
                        segments.Add(new ModelSegment(start, end, areaName));
                    }
                }
            }

            return segments;
        }

        private double GetDefaultSearchRadius()
        {
            OperationResult<ExcelCSIToolBox.Data.DTOs.CSI.CSISapModelPresentUnitSystemDTO> unitResult =
                _csiConnectionService.GetPresentUnitSystem();
            int lengthUnit = unitResult.IsSuccess && unitResult.Data != null
                ? unitResult.Data.LengthUnit
                : (SelectedUnitSystem == null ? 6 : SelectedUnitSystem.LengthUnit);

            switch (lengthUnit)
            {
                case 1: return 393.7007874015748; // inch
                case 2: return 32.80839895013123; // ft
                case 4: return 10000.0; // mm
                case 5: return 1000.0; // cm
                case 6: return 10.0; // m
                default: return 10.0;
            }
        }

        private static double GetGeometryTolerance(double searchRadius)
        {
            return Math.Max(0.000001, searchRadius * 0.000001);
        }

        private static WorkingPlane DetectWorkingPlane(Vector3 candidateDirection, Vector3 arrayDirection)
        {
            Vector3 normal = Cross(candidateDirection, arrayDirection);
            double absX = Math.Abs(normal.X);
            double absY = Math.Abs(normal.Y);
            double absZ = Math.Abs(normal.Z);

            if (absZ >= absX && absZ >= absY)
            {
                return WorkingPlane.XY;
            }

            return absY >= absX ? WorkingPlane.XZ : WorkingPlane.YZ;
        }

        private static bool IsInSameWorkingPlane(
            Vector3 start,
            Vector3 end,
            WorkingPlane workingPlane,
            double constantCoordinate,
            double tolerance)
        {
            return Math.Abs(GetPlaneConstant(start, workingPlane) - constantCoordinate) <= tolerance &&
                   Math.Abs(GetPlaneConstant(end, workingPlane) - constantCoordinate) <= tolerance;
        }

        private static double GetPlaneConstant(Vector3 point, WorkingPlane workingPlane)
        {
            switch (workingPlane)
            {
                case WorkingPlane.XY: return point.Z;
                case WorkingPlane.XZ: return point.Y;
                default: return point.X;
            }
        }

        private static Point2D Project(Vector3 point, WorkingPlane workingPlane)
        {
            switch (workingPlane)
            {
                case WorkingPlane.XY: return new Point2D(point.X, point.Y);
                case WorkingPlane.XZ: return new Point2D(point.X, point.Z);
                default: return new Point2D(point.Y, point.Z);
            }
        }

        private static Vector3 Unproject(Point2D point, WorkingPlane workingPlane, double constantCoordinate)
        {
            switch (workingPlane)
            {
                case WorkingPlane.XY: return new Vector3(point.U, point.V, constantCoordinate);
                case WorkingPlane.XZ: return new Vector3(point.U, constantCoordinate, point.V);
                default: return new Vector3(constantCoordinate, point.U, point.V);
            }
        }

        private static Vector2D Subtract(Point2D left, Point2D right)
        {
            return new Vector2D(left.U - right.U, left.V - right.V);
        }

        private static double Dot(Vector2D left, Vector2D right)
        {
            return left.U * right.U + left.V * right.V;
        }

        private static double Cross(Vector2D left, Vector2D right)
        {
            return left.U * right.V - left.V * right.U;
        }

        private static Vector3 Cross(Vector3 left, Vector3 right)
        {
            return new Vector3(
                left.Y * right.Z - left.Z * right.Y,
                left.Z * right.X - left.X * right.Z,
                left.X * right.Y - left.Y * right.X);
        }

        private static double Distance(Vector3 left, Vector3 right)
        {
            return Length(Subtract(left, right));
        }

        private static string FormatCoordinate(double value)
        {
            return value.ToString("G10", CultureInfo.CurrentCulture);
        }

        private void ClearPoint1Coordinates()
        {
            Point1X = string.Empty;
            Point1Y = string.Empty;
            Point1Z = string.Empty;
        }

        private void ClearPoint2Coordinates()
        {
            Point2X = string.Empty;
            Point2Y = string.Empty;
            Point2Z = string.Empty;
        }

        private void CloseWindow(Window window)
        {
            if (window != null)
            {
                window.Close();
            }
        }

        private void ShowWarning(string message)
        {
            MessageBox.Show(
                string.IsNullOrWhiteSpace(message) ? "The requested operation could not be completed." : message,
                ProductTitle,
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
        }

        private sealed class ArrayFrameInput
        {
            public Vector3 Point1 { get; set; }
            public Vector3 Point2 { get; set; }
            public Vector3 ReferencePointI { get; set; }
            public Vector3 ReferencePointJ { get; set; }
        }

        private sealed class ArrayBetweenTwoLinesInput
        {
            public Vector3 Line1Start { get; set; }
            public Vector3 Line1End { get; set; }
            public Vector3 Line2Start { get; set; }
            public Vector3 Line2End { get; set; }
        }

        private enum WorkingPlane
        {
            XY,
            XZ,
            YZ
        }

        private sealed class ModelSegment
        {
            public ModelSegment(Vector3 start, Vector3 end, string sourceName)
            {
                Start = start;
                End = end;
                SourceName = sourceName;
                Bounds = BoundingBox.FromSegment(start, end, 0);
            }

            public Vector3 Start { get; }
            public Vector3 End { get; }
            public string SourceName { get; }
            public BoundingBox Bounds { get; }
        }

        private sealed class BoundingBox
        {
            private BoundingBox(
                double minX,
                double minY,
                double minZ,
                double maxX,
                double maxY,
                double maxZ)
            {
                MinX = minX;
                MinY = minY;
                MinZ = minZ;
                MaxX = maxX;
                MaxY = maxY;
                MaxZ = maxZ;
            }

            private double MinX { get; }
            private double MinY { get; }
            private double MinZ { get; }
            private double MaxX { get; }
            private double MaxY { get; }
            private double MaxZ { get; }

            public static BoundingBox FromSegment(Vector3 start, Vector3 end, double expansion)
            {
                return new BoundingBox(
                    Math.Min(start.X, end.X) - expansion,
                    Math.Min(start.Y, end.Y) - expansion,
                    Math.Min(start.Z, end.Z) - expansion,
                    Math.Max(start.X, end.X) + expansion,
                    Math.Max(start.Y, end.Y) + expansion,
                    Math.Max(start.Z, end.Z) + expansion);
            }

            public bool Intersects(BoundingBox other)
            {
                return other != null &&
                       MinX <= other.MaxX &&
                       MaxX >= other.MinX &&
                       MinY <= other.MaxY &&
                       MaxY >= other.MinY &&
                       MinZ <= other.MaxZ &&
                       MaxZ >= other.MinZ;
            }
        }

        private sealed class IntersectionResult
        {
            public Vector3 Point { get; set; }
            public double CandidateParameter { get; set; }
        }

        private sealed class Intersection2DResult
        {
            public Point2D Point { get; set; }
            public double CandidateParameter { get; set; }
        }

        private sealed class CandidateAdjustmentResult
        {
            public Vector3 Start { get; set; }
            public Vector3 End { get; set; }
            public bool IsSkipped { get; set; }
            public bool IsTrimmed { get; set; }
            public bool IsExtended { get; set; }

            public static CandidateAdjustmentResult Skipped(Vector3 start, Vector3 end)
            {
                return new CandidateAdjustmentResult
                {
                    Start = start,
                    End = end,
                    IsSkipped = true
                };
            }
        }

        private sealed class TrimExtendSummary
        {
            public int AdjustedCount { get; private set; }
            public int TrimmedCount { get; private set; }
            public int ExtendedCount { get; private set; }
            public int SkippedCount { get; private set; }

            public void Add(CandidateAdjustmentResult result)
            {
                if (result == null)
                {
                    return;
                }

                if (result.IsSkipped)
                {
                    SkippedCount++;
                    return;
                }

                if (result.IsTrimmed || result.IsExtended)
                {
                    AdjustedCount++;
                }

                if (result.IsTrimmed)
                {
                    TrimmedCount++;
                }

                if (result.IsExtended)
                {
                    ExtendedCount++;
                }
            }
        }

        private struct Point2D
        {
            public Point2D(double u, double v)
            {
                U = u;
                V = v;
            }

            public double U { get; }
            public double V { get; }
        }

        private struct Vector2D
        {
            public Vector2D(double u, double v)
            {
                U = u;
                V = v;
            }

            public double U { get; }
            public double V { get; }
        }

        private struct Vector3
        {
            public Vector3(double x, double y, double z)
            {
                X = x;
                Y = y;
                Z = z;
            }

            public double X { get; }
            public double Y { get; }
            public double Z { get; }
        }
    }
}
