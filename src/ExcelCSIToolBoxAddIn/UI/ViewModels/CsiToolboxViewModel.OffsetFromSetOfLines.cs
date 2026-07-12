using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Windows;
using System.Windows.Input;
using ExcelCSIToolBox.Application.Modelling.OffsetPolylines;
using ExcelCSIToolBox.Core.Abstractions.CSI;
using ExcelCSIToolBoxAddIn.UI.Common.Commands;
using ExcelCSIToolBox.Core.Common.Results;
using ExcelCSIToolBox.Core.Contracts.CSI.FrameObject;
using ExcelCSIToolBox.Core.Contracts.CSI.PointObject;
using ExcelCSIToolBox.Core.Contracts.CSI;

namespace ExcelCSIToolBoxAddIn.UI.ViewModels
{
    public partial class CsiToolboxViewModel
    {
        private OffsetPolylineService _offsetPolylineService;
        private OffsetPolylineResult _offsetValidationResult;
        private OffsetPolylineResult _offsetPreviewResult;
        private string _offsetDistanceText;
        private string _offsetCoordinateToleranceText;
        private string _offsetMiterLimitText;
        private string _offsetValidationStatus;
        private string _offsetPolygonOrientation;
        private string _offsetPlaneInfo;
        private string _offsetDirectionText;
        private string _offsetResultTypeText;
        private string _offsetStatusMessage;
        private string _offsetGroupName;
        private string _offsetSelectedSection;
        private bool _offsetCopySectionProperty;
        private bool _offsetAddResultToGroup;
        private bool _offsetSelectionIsValid;
        private int _offsetPreviewVersion;
        private ExcelCSIToolBoxAddIn.UI.Views.OffsetFromSetOfLinesWindow _offsetFromSetOfLinesWindow;

        private void InitializeOffsetFromSetOfLinesPage()
        {
            _offsetPolylineService = new OffsetPolylineService();
            OffsetSelectedLineSegments = new ObservableCollection<SourceLineSegment>();
            OffsetOrderedSegments = new ObservableCollection<OrderedLineSegment>();
            OffsetResultSegments = new ObservableCollection<OffsetLineSegment>();
            OffsetAvailableSections = new ObservableCollection<string>();

            _offsetDistanceText = string.Empty;
            _offsetCoordinateToleranceText = "0.001";
            _offsetMiterLimitText = "10";
            _offsetGroupName = "OffsetPolyline_01";
            _offsetCopySectionProperty = true;
            _offsetValidationStatus = "No line selection.";
            _offsetPolygonOrientation = "-";
            _offsetPlaneInfo = "-";
            _offsetDirectionText = "-";
            _offsetResultTypeText = "-";
            _offsetStatusMessage = "Attach to ETABS, select frame objects forming one closed boundary, then get selected lines.";

            OpenOffsetFromSetOfLinesCommand = new RelayCommand(OpenOffsetFromSetOfLinesWindow);
            OffsetGetSelectedLinesCommand = new RelayCommand(GetOffsetSelectedLines, CanExecuteCsiAction);
            OffsetPreviewCommand = new RelayCommand(PreviewOffsetFromSetOfLines, CanPreviewOffsetFromSetOfLines);
            OffsetCreateInEtabsCommand = new RelayCommand(CreateOffsetFromSetOfLinesInEtabs, CanCreateOffsetFromSetOfLines);
            OffsetClearCommand = new RelayCommand(ClearOffsetFromSetOfLines);
            OffsetRefreshSectionsCommand = new RelayCommand(RefreshAvailableFrameSectionsForOffsetTool, CanExecuteCsiAction);
        }

        public ObservableCollection<SourceLineSegment> OffsetSelectedLineSegments { get; private set; }
        public ObservableCollection<OrderedLineSegment> OffsetOrderedSegments { get; private set; }
        public ObservableCollection<OffsetLineSegment> OffsetResultSegments { get; private set; }
        public ObservableCollection<string> OffsetAvailableSections { get; private set; }

        public ICommand OpenOffsetFromSetOfLinesCommand { get; private set; }
        public ICommand OffsetGetSelectedLinesCommand { get; private set; }
        public ICommand OffsetPreviewCommand { get; private set; }
        public ICommand OffsetCreateInEtabsCommand { get; private set; }
        public ICommand OffsetClearCommand { get; private set; }
        public ICommand OffsetRefreshSectionsCommand { get; private set; }

        public int OffsetSelectedLineCount
        {
            get { return OffsetSelectedLineSegments == null ? 0 : OffsetSelectedLineSegments.Count; }
        }

        public int OffsetDetectedVertexCount
        {
            get { return _offsetValidationResult == null ? 0 : _offsetValidationResult.DetectedVertexCount; }
        }

        public int OffsetOrderedSegmentCount
        {
            get { return OffsetOrderedSegments == null ? 0 : OffsetOrderedSegments.Count; }
        }

        public int OffsetResultSegmentCount
        {
            get { return OffsetResultSegments == null ? 0 : OffsetResultSegments.Count; }
        }

        public string OffsetDistanceText
        {
            get { return _offsetDistanceText; }
            set
            {
                if (_offsetDistanceText == value) return;
                _offsetDistanceText = value;
                OnPropertyChanged();
                InvalidateOffsetPreview("Offset distance changed. Preview must be recalculated.");
            }
        }

        public string OffsetCoordinateToleranceText
        {
            get { return _offsetCoordinateToleranceText; }
            set
            {
                if (_offsetCoordinateToleranceText == value) return;
                _offsetCoordinateToleranceText = value;
                OnPropertyChanged();
                RevalidateOffsetSelection();
                InvalidateOffsetPreview("Coordinate tolerance changed. Preview must be recalculated.");
            }
        }

        public string OffsetMiterLimitText
        {
            get { return _offsetMiterLimitText; }
            set
            {
                if (_offsetMiterLimitText == value) return;
                _offsetMiterLimitText = value;
                OnPropertyChanged();
                InvalidateOffsetPreview("Miter limit changed. Preview must be recalculated.");
            }
        }

        public bool OffsetCopySectionProperty
        {
            get { return _offsetCopySectionProperty; }
            set
            {
                if (_offsetCopySectionProperty == value) return;
                _offsetCopySectionProperty = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(OffsetAssignedSectionText));
                OnPropertyChanged(nameof(OffsetAssignedSectionInputEnabled));
                InvalidateOffsetPreview("Output options changed. Preview must be recalculated.");
            }
        }

        public bool OffsetAssignedSectionInputEnabled
        {
            get { return !OffsetCopySectionProperty; }
        }

        public string OffsetSelectedSection
        {
            get { return _offsetSelectedSection; }
            set
            {
                if (_offsetSelectedSection == value) return;
                _offsetSelectedSection = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(OffsetAssignedSectionText));
                if (!OffsetCopySectionProperty)
                {
                    InvalidateOffsetPreview("Section assignment changed. Preview must be recalculated.");
                }
                else
                {
                    RefreshOffsetCommandStates();
                }
            }
        }

        public bool OffsetAddResultToGroup
        {
            get { return _offsetAddResultToGroup; }
            set
            {
                if (_offsetAddResultToGroup == value) return;
                _offsetAddResultToGroup = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(OffsetGroupNameInputEnabled));
                InvalidateOffsetPreview("Output options changed. Preview must be recalculated.");
            }
        }

        public bool OffsetGroupNameInputEnabled
        {
            get { return OffsetAddResultToGroup; }
        }

        public string OffsetGroupName
        {
            get { return _offsetGroupName; }
            set
            {
                if (_offsetGroupName == value) return;
                _offsetGroupName = value;
                OnPropertyChanged();
                InvalidateOffsetPreview("Group option changed. Preview must be recalculated.");
            }
        }

        public string OffsetValidationStatus
        {
            get { return _offsetValidationStatus; }
            private set
            {
                if (_offsetValidationStatus == value) return;
                _offsetValidationStatus = value;
                OnPropertyChanged();
            }
        }

        public string OffsetPolygonOrientation
        {
            get { return _offsetPolygonOrientation; }
            private set
            {
                if (_offsetPolygonOrientation == value) return;
                _offsetPolygonOrientation = value;
                OnPropertyChanged();
            }
        }

        public string OffsetPlaneInfo
        {
            get { return _offsetPlaneInfo; }
            private set
            {
                if (_offsetPlaneInfo == value) return;
                _offsetPlaneInfo = value;
                OnPropertyChanged();
            }
        }

        public string OffsetDirectionText
        {
            get { return _offsetDirectionText; }
            private set
            {
                if (_offsetDirectionText == value) return;
                _offsetDirectionText = value;
                OnPropertyChanged();
            }
        }

        public string OffsetResultTypeText
        {
            get { return _offsetResultTypeText; }
            private set
            {
                if (_offsetResultTypeText == value) return;
                _offsetResultTypeText = value;
                OnPropertyChanged();
            }
        }

        public string OffsetStatusMessage
        {
            get { return _offsetStatusMessage; }
            private set
            {
                if (_offsetStatusMessage == value) return;
                _offsetStatusMessage = value;
                OnPropertyChanged();
            }
        }

        public OffsetPolylineResult OffsetPreviewResult
        {
            get { return _offsetPreviewResult; }
            private set
            {
                if (ReferenceEquals(_offsetPreviewResult, value)) return;
                _offsetPreviewResult = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(CanCreateOffsetPreview));
                IncrementOffsetPreviewVersion();
                RefreshOffsetCommandStates();
            }
        }

        public int OffsetPreviewVersion
        {
            get { return _offsetPreviewVersion; }
            private set
            {
                if (_offsetPreviewVersion == value) return;
                _offsetPreviewVersion = value;
                OnPropertyChanged();
            }
        }

        public string OffsetLengthUnitText
        {
            get
            {
                string text = GetLengthUnitText();
                return string.IsNullOrWhiteSpace(text) ? "model length unit" : text;
            }
        }

        public string OffsetAssignedSectionText
        {
            get
            {
                if (OffsetCopySectionProperty)
                {
                    return "Source segment section property";
                }

                return string.IsNullOrWhiteSpace(OffsetSelectedSection)
                    ? "Select an ETABS frame section."
                    : "Assigned section: " + OffsetSelectedSection;
            }
        }

        public bool CanCreateOffsetPreview
        {
            get { return OffsetPreviewResult != null && OffsetPreviewResult.ResultSegments != null && OffsetPreviewResult.ResultSegments.Count > 0; }
        }

        private void OpenOffsetFromSetOfLinesWindow()
        {
            RefreshAvailableFrameSectionsForOffsetTool();

            if (_offsetFromSetOfLinesWindow != null)
            {
                _offsetFromSetOfLinesWindow.Activate();
                return;
            }

            var window = new ExcelCSIToolBoxAddIn.UI.Views.OffsetFromSetOfLinesWindow(this);
            Window owner = GetActiveOwnerWindow();
            if (owner != null && !ReferenceEquals(owner, window))
            {
                window.Owner = owner;
                window.WindowStartupLocation = WindowStartupLocation.CenterOwner;
            }

            window.Closed += delegate { _offsetFromSetOfLinesWindow = null; };
            _offsetFromSetOfLinesWindow = window;
            window.Show();
        }

        private void GetOffsetSelectedLines()
        {
            ClearOffsetCollections();
            _offsetValidationResult = null;
            _offsetSelectionIsValid = false;
            OffsetPreviewResult = null;

            if (!CanUseActiveModel)
            {
                OffsetValidationStatus = "No ETABS model is attached.";
                OffsetStatusMessage = "Attach to a running ETABS model before reading selected lines.";
                RefreshOffsetCommandStates();
                return;
            }

            OperationResult<IReadOnlyList<CsiSelectedObjectDto>> selectedResult =
                _csiConnectionService.GetSelectedObjectsFromActiveModel();
            if (!selectedResult.IsSuccess)
            {
                OffsetValidationStatus = selectedResult.Message;
                OffsetStatusMessage = selectedResult.Message;
                RefreshOffsetCountsAndCommands();
                return;
            }

            List<string> frameNames = ExtractSelectedFrameNames(selectedResult.Data);
            if (frameNames.Count == 0)
            {
                OffsetValidationStatus = "At least three line objects must be selected.";
                OffsetStatusMessage = "Select ETABS frame objects that collectively form one closed boundary.";
                RefreshOffsetCountsAndCommands();
                return;
            }

            OperationResult<List<SourceLineSegment>> sourceResult = ReadSourceLineSegments(frameNames);
            if (!sourceResult.IsSuccess)
            {
                OffsetValidationStatus = sourceResult.Message;
                OffsetStatusMessage = sourceResult.Message;
                RefreshOffsetCountsAndCommands();
                return;
            }

            foreach (SourceLineSegment segment in sourceResult.Data)
            {
                OffsetSelectedLineSegments.Add(segment);
            }

            RevalidateOffsetSelection();
            if (_offsetSelectionIsValid)
            {
                OffsetStatusMessage = "Selection is valid. Enter a non-zero offset distance and preview the result.";
            }

            RefreshOffsetCountsAndCommands();
        }

        private void PreviewOffsetFromSetOfLines()
        {
            OffsetPolylineOptions options;
            double offsetDistance;
            OperationResult inputResult = TryReadOffsetInputs(out offsetDistance, out options);
            if (!inputResult.IsSuccess)
            {
                OffsetStatusMessage = inputResult.Message;
                RefreshOffsetCommandStates();
                return;
            }

            if (!_offsetSelectionIsValid)
            {
                RevalidateOffsetSelection();
            }

            if (!_offsetSelectionIsValid)
            {
                OffsetStatusMessage = OffsetValidationStatus;
                RefreshOffsetCommandStates();
                return;
            }

            string groupName = OffsetAddResultToGroup ? CleanGroupName(OffsetGroupName) : null;
            if (OffsetAddResultToGroup && string.IsNullOrWhiteSpace(groupName))
            {
                OffsetStatusMessage = "Group name is required.";
                RefreshOffsetCommandStates();
                return;
            }

            OperationResult<OffsetPolylineResult> result =
                _offsetPolylineService.CalculateOffset(
                    new List<SourceLineSegment>(OffsetSelectedLineSegments),
                    offsetDistance,
                    options,
                    groupName);

            if (!result.IsSuccess)
            {
                OffsetPreviewResult = null;
                OffsetResultSegments.Clear();
                OffsetResultTypeText = "-";
                OffsetDirectionText = offsetDistance > 0 ? "Outward" : "Inward";
                OffsetStatusMessage = result.Message;
                RefreshOffsetCountsAndCommands();
                return;
            }

            OffsetResultSegments.Clear();
            foreach (OffsetLineSegment segment in result.Data.ResultSegments)
            {
                OffsetResultSegments.Add(segment);
            }

            OffsetPreviewResult = result.Data;
            OffsetDirectionText = result.Data.OffsetDirection;
            OffsetResultTypeText = result.Data.ResultType;
            OffsetStatusMessage = result.Data.ValidationMessage;
            RefreshOffsetCountsAndCommands();
        }

        private void CreateOffsetFromSetOfLinesInEtabs()
        {
            if (OffsetPreviewResult == null || OffsetPreviewResult.ResultSegments == null || OffsetPreviewResult.ResultSegments.Count == 0)
            {
                OffsetStatusMessage = "Preview the offset before creating ETABS objects.";
                return;
            }

            if (!CanUseActiveModel)
            {
                OffsetStatusMessage = "No ETABS model is attached.";
                return;
            }

            if (!TryValidateOffsetSectionAssignment())
            {
                return;
            }

            if (!ConfirmSourceSelectionStillMatches())
            {
                return;
            }

            OperationResult<bool> lockResult = _csiConnectionService.GetModelIsLocked();
            if (lockResult.IsSuccess && lockResult.Data)
            {
                OffsetStatusMessage = "The ETABS model is locked. Unlock the model before creating objects.";
                return;
            }

            var frameRequests = new List<FrameAddRequestDto>();
            foreach (OffsetLineSegment segment in OffsetPreviewResult.ResultSegments)
            {
                frameRequests.Add(new FrameAddRequestDto
                {
                    SectionName = GetOffsetResultSectionName(segment),
                    Xi = segment.StartX,
                    Yi = segment.StartY,
                    Zi = segment.StartZ,
                    Xj = segment.EndX,
                    Yj = segment.EndY,
                    Zj = segment.EndZ
                });
            }

            OperationResult<FrameAddBatchResultDto> addResult =
                _csiConnectionService.AddFrameObjects(new FrameAddBatchRequestDto
                {
                    Frames = frameRequests,
                    SuppressViewRefresh = true
                });

            if (!addResult.IsSuccess || addResult.Data == null)
            {
                OffsetStatusMessage = string.IsNullOrWhiteSpace(addResult.Message)
                    ? "ETABS frame creation failed."
                    : addResult.Message;
                return;
            }

            FrameAddBatchResultDto batch = addResult.Data;
            List<string> createdNames = batch.SuccessfulFrameNames ?? new List<string>();
            if (batch.FailureCount > 0 || createdNames.Count != OffsetPreviewResult.ResultSegments.Count)
            {
                RollBackCreatedOffsetFrames(createdNames);
                OffsetStatusMessage = BuildBatchFailureMessage(batch);
                return;
            }

            ApplyCreatedNamesToResultSegments(createdNames);

            if (OffsetAddResultToGroup)
            {
                OperationResult groupResult = _csiConnectionService.AddFramesToGroup(createdNames, CleanGroupName(OffsetGroupName));
                if (!groupResult.IsSuccess)
                {
                    RollBackCreatedOffsetFrames(createdNames);
                    OffsetStatusMessage = "Creation rolled back because group assignment failed: " + groupResult.Message;
                    return;
                }
            }

            OperationResult selectResult = _csiConnectionService.SelectFramesByUniqueNames(createdNames);
            _csiConnectionService.RefreshView(false);

            OffsetStatusMessage = selectResult.IsSuccess
                ? "Created " + createdNames.Count.ToString(CultureInfo.CurrentCulture) + " offset frame object(s)."
                : "Created " + createdNames.Count.ToString(CultureInfo.CurrentCulture) + " offset frame object(s), but selection failed: " + selectResult.Message;
            StatusText = OffsetStatusMessage;
            RefreshOffsetCountsAndCommands();
        }

        private void ClearOffsetFromSetOfLines()
        {
            ClearOffsetCollections();
            _offsetValidationResult = null;
            _offsetSelectionIsValid = false;
            OffsetValidationStatus = "No line selection.";
            OffsetPolygonOrientation = "-";
            OffsetPlaneInfo = "-";
            OffsetDirectionText = "-";
            OffsetResultTypeText = "-";
            OffsetStatusMessage = "Offset from Set of Lines was cleared.";
            OffsetPreviewResult = null;
            RefreshOffsetCountsAndCommands();
        }

        private void RevalidateOffsetSelection()
        {
            _offsetValidationResult = null;
            _offsetSelectionIsValid = false;
            OffsetOrderedSegments.Clear();
            OffsetPolygonOrientation = "-";
            OffsetPlaneInfo = "-";

            if (OffsetSelectedLineSegments.Count == 0)
            {
                OffsetValidationStatus = "No line selection.";
                RefreshOffsetCountsAndCommands();
                return;
            }

            OffsetPolylineOptions options;
            double ignoredOffset;
            OperationResult inputResult = TryReadOffsetInputs(out ignoredOffset, out options, allowEmptyOffset: true);
            if (!inputResult.IsSuccess)
            {
                OffsetValidationStatus = inputResult.Message;
                RefreshOffsetCountsAndCommands();
                return;
            }

            OperationResult<OffsetPolylineResult> validationResult =
                _offsetPolylineService.ValidateClosedBoundary(
                    new List<SourceLineSegment>(OffsetSelectedLineSegments),
                    options);

            if (!validationResult.IsSuccess)
            {
                OffsetValidationStatus = validationResult.Message;
                RefreshOffsetCountsAndCommands();
                return;
            }

            _offsetValidationResult = validationResult.Data;
            _offsetSelectionIsValid = true;
            foreach (OrderedLineSegment segment in validationResult.Data.OrderedSegments)
            {
                OffsetOrderedSegments.Add(segment);
            }

            OffsetValidationStatus = validationResult.Data.ValidationMessage;
            OffsetPolygonOrientation = validationResult.Data.PolygonOrientation;
            OffsetPlaneInfo = FormatPlaneInfo(validationResult.Data);
            RefreshOffsetCountsAndCommands();
        }

        private OperationResult<List<SourceLineSegment>> ReadSourceLineSegments(IReadOnlyList<string> frameNames)
        {
            var segments = new List<SourceLineSegment>();
            for (int i = 0; i < frameNames.Count; i++)
            {
                string frameName = frameNames[i];
                OperationResult<FrameEndPointInfo> pointsResult = _csiConnectionService.GetFramePoints(frameName);
                if (!pointsResult.IsSuccess || pointsResult.Data == null)
                {
                    return OperationResult<List<SourceLineSegment>>.Failure(pointsResult.Message);
                }

                OperationResult<PointObjectInfo> pointIResult = _csiConnectionService.GetPointCoordinates(pointsResult.Data.PointI);
                OperationResult<PointObjectInfo> pointJResult = _csiConnectionService.GetPointCoordinates(pointsResult.Data.PointJ);
                if (!pointIResult.IsSuccess || pointIResult.Data == null)
                {
                    return OperationResult<List<SourceLineSegment>>.Failure(pointIResult.Message);
                }

                if (!pointJResult.IsSuccess || pointJResult.Data == null)
                {
                    return OperationResult<List<SourceLineSegment>>.Failure(pointJResult.Message);
                }

                OperationResult<FrameSectionInfo> sectionResult = _csiConnectionService.GetFrameSection(frameName);
                string sectionName = sectionResult.IsSuccess && sectionResult.Data != null
                    ? sectionResult.Data.SectionName
                    : string.Empty;

                double length = Distance(pointIResult.Data, pointJResult.Data);
                segments.Add(new SourceLineSegment
                {
                    ObjectName = frameName,
                    SelectionIndex = i + 1,
                    IPointName = pointsResult.Data.PointI,
                    JPointName = pointsResult.Data.PointJ,
                    StartX = pointIResult.Data.X,
                    StartY = pointIResult.Data.Y,
                    StartZ = pointIResult.Data.Z,
                    EndX = pointJResult.Data.X,
                    EndY = pointJResult.Data.Y,
                    EndZ = pointJResult.Data.Z,
                    SectionProperty = sectionName,
                    StoryName = FormatStoryOrElevation(pointIResult.Data, pointJResult.Data),
                    Length = length
                });
            }

            return OperationResult<List<SourceLineSegment>>.Success(segments);
        }

        private bool ConfirmSourceSelectionStillMatches()
        {
            OperationResult<IReadOnlyList<CsiSelectedObjectDto>> selectedResult =
                _csiConnectionService.GetSelectedObjectsFromActiveModel();
            if (!selectedResult.IsSuccess)
            {
                OffsetStatusMessage = selectedResult.Message;
                return false;
            }

            List<string> currentFrameNames = ExtractSelectedFrameNames(selectedResult.Data);
            var current = new HashSet<string>(currentFrameNames, StringComparer.OrdinalIgnoreCase);
            var source = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (SourceLineSegment segment in OffsetSelectedLineSegments)
            {
                if (!string.IsNullOrWhiteSpace(segment.ObjectName))
                {
                    source.Add(segment.ObjectName);
                }
            }

            if (current.SetEquals(source))
            {
                return true;
            }

            InvalidateOffsetPreview("The selected ETABS objects changed. Get selected lines and preview again.");
            OffsetStatusMessage = "The selected ETABS objects changed. Get selected lines and preview again.";
            return false;
        }

        private OperationResult TryReadOffsetInputs(
            out double offsetDistance,
            out OffsetPolylineOptions options,
            bool allowEmptyOffset = false)
        {
            offsetDistance = 0;
            options = null;

            double coordinateTolerance;
            if (!TryParseDouble(OffsetCoordinateToleranceText, out coordinateTolerance) || coordinateTolerance <= 0)
            {
                return OperationResult.Failure("Coordinate tolerance must be a positive number.");
            }

            double miterLimit;
            if (!TryParseDouble(OffsetMiterLimitText, out miterLimit) || miterLimit < 1)
            {
                return OperationResult.Failure("Miter limit must be a number greater than or equal to 1.");
            }

            if (string.IsNullOrWhiteSpace(OffsetDistanceText) && allowEmptyOffset)
            {
                offsetDistance = 0;
            }
            else if (!TryParseDouble(OffsetDistanceText, out offsetDistance))
            {
                return OperationResult.Failure("Offset distance must be a valid number.");
            }

            if (!allowEmptyOffset && Math.Abs(offsetDistance) <= coordinateTolerance)
            {
                return OperationResult.Failure("Offset distance cannot be zero.");
            }

            options = new OffsetPolylineOptions
            {
                CoordinateTolerance = coordinateTolerance,
                PlaneTolerance = coordinateTolerance,
                ZeroLengthTolerance = coordinateTolerance,
                ParallelTolerance = Math.Max(0.000000001, coordinateTolerance * 0.000001),
                AreaTolerance = Math.Max(0.000000001, coordinateTolerance * coordinateTolerance),
                MiterLimit = miterLimit
            };

            return OperationResult.Success();
        }

        private bool CanPreviewOffsetFromSetOfLines()
        {
            OffsetPolylineOptions options;
            double offsetDistance;
            return CanUseActiveModel &&
                   _offsetSelectionIsValid &&
                   TryReadOffsetInputs(out offsetDistance, out options).IsSuccess;
        }

        private bool CanCreateOffsetFromSetOfLines()
        {
            return CanUseActiveModel &&
                   OffsetPreviewResult != null &&
                   OffsetPreviewResult.ResultSegments != null &&
                   OffsetPreviewResult.ResultSegments.Count > 0 &&
                   HasOffsetSectionAssignment();
        }

        private bool TryValidateOffsetSectionAssignment()
        {
            if (HasOffsetSectionAssignment())
            {
                return true;
            }

            OffsetStatusMessage = "Select an ETABS frame section or enable Copy Section Property.";
            RefreshOffsetCommandStates();
            return false;
        }

        private bool HasOffsetSectionAssignment()
        {
            return OffsetCopySectionProperty || !string.IsNullOrWhiteSpace(OffsetSelectedSection);
        }

        private string GetOffsetResultSectionName(OffsetLineSegment segment)
        {
            if (OffsetCopySectionProperty)
            {
                return segment == null ? string.Empty : segment.SourceSectionProperty;
            }

            return string.IsNullOrWhiteSpace(OffsetSelectedSection)
                ? string.Empty
                : OffsetSelectedSection;
        }

        private void RefreshAvailableFrameSectionsForOffsetTool()
        {
            OffsetAvailableSections.Clear();

            if (!IsConnected)
            {
                OffsetSelectedSection = null;
                return;
            }

            OperationResult<IReadOnlyList<string>> result = _csiConnectionService.GetFrameSectionNames();
            if (!result.IsSuccess || result.Data == null)
            {
                ShowWarning(result.Message);
                return;
            }

            string selectedSection = OffsetSelectedSection;
            foreach (string sectionName in result.Data)
            {
                if (!string.IsNullOrWhiteSpace(sectionName))
                {
                    OffsetAvailableSections.Add(sectionName);
                }
            }

            if (!string.IsNullOrWhiteSpace(selectedSection))
            {
                foreach (string sectionName in OffsetAvailableSections)
                {
                    if (string.Equals(sectionName, selectedSection, StringComparison.OrdinalIgnoreCase))
                    {
                        OffsetSelectedSection = sectionName;
                        return;
                    }
                }
            }

            OffsetSelectedSection = OffsetAvailableSections.Count > 0
                ? OffsetAvailableSections[0]
                : null;
        }

        private void InvalidateOffsetPreview(string message)
        {
            if (OffsetResultSegments != null)
            {
                OffsetResultSegments.Clear();
            }

            _offsetPreviewResult = null;
            OnPropertyChanged(nameof(OffsetPreviewResult));
            OnPropertyChanged(nameof(CanCreateOffsetPreview));
            OffsetDirectionText = "-";
            OffsetResultTypeText = "-";
            if (!string.IsNullOrWhiteSpace(message))
            {
                OffsetStatusMessage = message;
            }

            IncrementOffsetPreviewVersion();
            RefreshOffsetCountsAndCommands();
        }

        private void ClearOffsetCollections()
        {
            OffsetSelectedLineSegments.Clear();
            OffsetOrderedSegments.Clear();
            OffsetResultSegments.Clear();
        }

        private void RefreshOffsetCountsAndCommands()
        {
            OnPropertyChanged(nameof(OffsetSelectedLineCount));
            OnPropertyChanged(nameof(OffsetDetectedVertexCount));
            OnPropertyChanged(nameof(OffsetOrderedSegmentCount));
            OnPropertyChanged(nameof(OffsetResultSegmentCount));
            OnPropertyChanged(nameof(CanCreateOffsetPreview));
            RefreshOffsetCommandStates();
        }

        private void RefreshOffsetCommandStates()
        {
            RaiseCommandState(OffsetGetSelectedLinesCommand);
            RaiseCommandState(OffsetPreviewCommand);
            RaiseCommandState(OffsetCreateInEtabsCommand);
            RaiseCommandState(OffsetClearCommand);
            RaiseCommandState(OffsetRefreshSectionsCommand);
        }

        private static void RaiseCommandState(ICommand command)
        {
            IRelayCommand relay = command as IRelayCommand;
            if (relay != null)
            {
                relay.RaiseCanExecuteChanged();
            }
        }

        private void IncrementOffsetPreviewVersion()
        {
            OffsetPreviewVersion = OffsetPreviewVersion + 1;
        }

        private void RollBackCreatedOffsetFrames(IReadOnlyList<string> createdNames)
        {
            if (createdNames == null || createdNames.Count == 0)
            {
                return;
            }

            _csiConnectionService.DeleteFrameObjects(createdNames);
            _csiConnectionService.RefreshView(false);
        }

        private void ApplyCreatedNamesToResultSegments(IReadOnlyList<string> createdNames)
        {
            int count = Math.Min(createdNames.Count, OffsetResultSegments.Count);
            for (int i = 0; i < count; i++)
            {
                OffsetResultSegments[i].NewObjectName = createdNames[i];
            }

            if (OffsetPreviewResult != null && OffsetPreviewResult.ResultSegments != null)
            {
                for (int i = 0; i < count && i < OffsetPreviewResult.ResultSegments.Count; i++)
                {
                    OffsetPreviewResult.ResultSegments[i].NewObjectName = createdNames[i];
                }
            }
        }

        private static string BuildBatchFailureMessage(FrameAddBatchResultDto batch)
        {
            if (batch == null)
            {
                return "ETABS frame creation failed.";
            }

            var parts = new List<string>();
            if (batch.FailedItems != null)
            {
                foreach (FrameAddResultDto item in batch.FailedItems)
                {
                    if (item != null && !string.IsNullOrWhiteSpace(item.FailureReason))
                    {
                        parts.Add(item.FailureReason);
                    }
                }
            }

            string detail = parts.Count == 0 ? string.Empty : " " + string.Join(" ", parts);
            return "Creation failed and all objects created by this operation were rolled back." + detail;
        }

        private static List<string> ExtractSelectedFrameNames(IReadOnlyList<CsiSelectedObjectDto> selectedObjects)
        {
            var frameNames = new List<string>();
            if (selectedObjects == null)
            {
                return frameNames;
            }

            foreach (CsiSelectedObjectDto selectedObject in selectedObjects)
            {
                if (selectedObject == null || string.IsNullOrWhiteSpace(selectedObject.UniqueName))
                {
                    continue;
                }

                if (string.Equals(selectedObject.ObjectType, "Frame", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(selectedObject.ObjectType, "Line", StringComparison.OrdinalIgnoreCase))
                {
                    frameNames.Add(selectedObject.UniqueName);
                }
            }

            return frameNames;
        }

        private static string FormatPlaneInfo(OffsetPolylineResult result)
        {
            if (result == null)
            {
                return "-";
            }

            return "Normal (" +
                   FormatNumber(result.PlaneNormal.X) + ", " +
                   FormatNumber(result.PlaneNormal.Y) + ", " +
                   FormatNumber(result.PlaneNormal.Z) + ")";
        }

        private static string FormatStoryOrElevation(PointObjectInfo pointI, PointObjectInfo pointJ)
        {
            if (pointI == null || pointJ == null)
            {
                return string.Empty;
            }

            return Math.Abs(pointI.Z - pointJ.Z) <= 0.000001
                ? "Z=" + FormatNumber(pointI.Z)
                : "Inclined";
        }

        private static double Distance(PointObjectInfo left, PointObjectInfo right)
        {
            double dx = left.X - right.X;
            double dy = left.Y - right.Y;
            double dz = left.Z - right.Z;
            return Math.Sqrt(dx * dx + dy * dy + dz * dz);
        }

        private static bool TryParseDouble(string text, out double value)
        {
            return double.TryParse(text, NumberStyles.Float, CultureInfo.CurrentCulture, out value) ||
                   double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out value);
        }

        private static string FormatNumber(double value)
        {
            return value.ToString("G8", CultureInfo.CurrentCulture);
        }

        private static string CleanGroupName(string groupName)
        {
            return string.IsNullOrWhiteSpace(groupName) ? string.Empty : groupName.Trim();
        }
    }
}
