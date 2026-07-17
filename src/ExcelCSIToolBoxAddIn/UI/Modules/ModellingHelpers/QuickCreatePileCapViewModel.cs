using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Text;
using System.Windows;
using System.Windows.Input;
using ExcelCSIToolBox.Application.Modelling.PileCaps;
using ExcelCSIToolBox.Core.Abstractions.CSI;
using ExcelCSIToolBox.Core.Common.Results;
using ExcelCSIToolBox.Core.Contracts.CSI;
using ExcelCSIToolBox.Core.Contracts.CSI.PileCap;
using ExcelCSIToolBoxAddIn.UI.Common.Commands;

namespace ExcelCSIToolBoxAddIn.UI.ViewModels
{
    public class QuickCreatePileCapViewModel : ViewModelBase
    {
        private readonly ICSISapModelConnectionService _connectionService;
        private readonly PileCapGeometryCalculator _geometryCalculator;
        private readonly PileCapInputValidator _validator;
        private readonly Action _closeWindow;
        private readonly Action _restoreFocus;
        private PileCapArrangementType _selectedArrangementType;
        private string _connectedModelName;
        private string _currentUnitsText;
        private int _selectedPointCount;
        private int _ignoredNonPointObjectCount;
        private string _selectedPointNamesText;
        private string _pileDiameterText;
        private string _pileLengthText;
        private string _rotationText;
        private bool _autoSpacing;
        private string _pileSpacingText;
        private string _spacingXText;
        private string _spacingYText;
        private bool _lockSpacingXAndY;
        private string _pileCapThicknessText;
        private string _edgeDistanceText;
        private string _selectedPileMaterial;
        private string _selectedPileCapMaterial;
        private string _pilePropertyName;
        private string _pileCapPropertyName;
        private string _validationMessage;
        private string _statusMessage;
        private bool _isProcessing;
        private bool _hasValidInputs;
        private PileCapGeometry _previewGeometry;
        private PileCapGeometry _monoCardGeometry;
        private PileCapGeometry _twoPileCardGeometry;
        private PileCapGeometry _threePileCardGeometry;
        private PileCapGeometry _fourPileCardGeometry;
        private int _previewVersion;
        private double _previewPileDiameter;
        private double _previewPileCapThickness;
        private double _previewEdgeDistance;
        private PileCapAssignmentSummaryDto _lastAssignmentSummary;

        public QuickCreatePileCapViewModel(
            ICSISapModelConnectionService connectionService,
            Action closeWindow = null,
            Action restoreFocus = null)
        {
            _connectionService = connectionService ?? throw new ArgumentNullException("connectionService");
            _geometryCalculator = new PileCapGeometryCalculator();
            _validator = new PileCapInputValidator();
            _closeWindow = closeWindow ?? delegate { };
            _restoreFocus = restoreFocus ?? delegate { };
            ConcreteMaterials = new ObservableCollection<string>();
            _selectedArrangementType = PileCapArrangementType.Mono;
            _pileDiameterText = "800";
            _pileLengthText = "30000";
            _rotationText = "0";
            _autoSpacing = true;
            _pileSpacingText = "2400";
            _spacingXText = "2400";
            _spacingYText = "2400";
            _lockSpacingXAndY = true;
            _pileCapThicknessText = "1500";
            _edgeDistanceText = "150";
            _connectedModelName = "-";
            _currentUnitsText = "-";
            _selectedPointNamesText = "-";
            _statusMessage = "Select point objects in ETABS, then click Assign.";

            RefreshSelectionCommand = new RelayCommand(RefreshSelectionAndRestoreFocus);
            AssignCommand = new RelayCommand(Assign, CanAssign);
            CloseCommand = new RelayCommand(delegate { _closeWindow(); });
            SelectArrangementCommand = new RelayCommand<object>(SelectArrangement);

            RefreshContext();
            RefreshGeneratedPropertiesAndPreview();
        }

        public ObservableCollection<string> ConcreteMaterials { get; private set; }

        public ICommand RefreshSelectionCommand { get; private set; }

        public ICommand AssignCommand { get; private set; }

        public ICommand CloseCommand { get; private set; }

        public ICommand SelectArrangementCommand { get; private set; }

        public string ConnectedModelName
        {
            get { return _connectedModelName; }
            private set
            {
                if (_connectedModelName == value) return;
                _connectedModelName = value;
                OnPropertyChanged();
            }
        }

        public string CurrentUnitsText
        {
            get { return _currentUnitsText; }
            private set
            {
                if (_currentUnitsText == value) return;
                _currentUnitsText = value;
                OnPropertyChanged();
            }
        }

        public int SelectedPointCount
        {
            get { return _selectedPointCount; }
            private set
            {
                if (_selectedPointCount == value) return;
                _selectedPointCount = value;
                OnPropertyChanged();
            }
        }

        public int IgnoredNonPointObjectCount
        {
            get { return _ignoredNonPointObjectCount; }
            private set
            {
                if (_ignoredNonPointObjectCount == value) return;
                _ignoredNonPointObjectCount = value;
                OnPropertyChanged();
            }
        }

        public string SelectedPointNamesText
        {
            get { return _selectedPointNamesText; }
            private set
            {
                if (_selectedPointNamesText == value) return;
                _selectedPointNamesText = value;
                OnPropertyChanged();
            }
        }

        public PileCapArrangementType SelectedArrangementType
        {
            get { return _selectedArrangementType; }
            set
            {
                if (_selectedArrangementType == value) return;
                _selectedArrangementType = value;
                OnPropertyChanged();
                OnPropertyChanged("SelectedArrangement");
                OnPropertyChanged("IsMonoSelected");
                OnPropertyChanged("IsTwoPileSelected");
                OnPropertyChanged("IsThreePileSelected");
                OnPropertyChanged("IsFourPileSelected");
                OnPropertyChanged("MonoSpacingVisibility");
                OnPropertyChanged("SingleSpacingVisibility");
                OnPropertyChanged("FourSpacingVisibility");
                OnPropertyChanged("SpacingDescriptionText");
                UpdateAutomaticSpacing();
                RefreshGeneratedPropertiesAndPreview();
            }
        }

        public PileCapArrangementType SelectedArrangement
        {
            get { return SelectedArrangementType; }
            set { SelectedArrangementType = value; }
        }

        public bool IsMonoSelected
        {
            get { return SelectedArrangementType == PileCapArrangementType.Mono; }
            set { if (value) SelectedArrangementType = PileCapArrangementType.Mono; }
        }

        public bool IsTwoPileSelected
        {
            get { return SelectedArrangementType == PileCapArrangementType.TwoPile; }
            set { if (value) SelectedArrangementType = PileCapArrangementType.TwoPile; }
        }

        public bool IsThreePileSelected
        {
            get { return SelectedArrangementType == PileCapArrangementType.ThreePile; }
            set { if (value) SelectedArrangementType = PileCapArrangementType.ThreePile; }
        }

        public bool IsFourPileSelected
        {
            get { return SelectedArrangementType == PileCapArrangementType.FourPile; }
            set { if (value) SelectedArrangementType = PileCapArrangementType.FourPile; }
        }

        public string PileDiameterText
        {
            get { return _pileDiameterText; }
            set
            {
                if (_pileDiameterText == value) return;
                _pileDiameterText = value;
                OnPropertyChanged();
                UpdateAutomaticSpacing();
                RefreshGeneratedPropertiesAndPreview();
            }
        }

        public string PileLengthText
        {
            get { return _pileLengthText; }
            set
            {
                if (_pileLengthText == value) return;
                _pileLengthText = value;
                OnPropertyChanged();
                RefreshGeneratedPropertiesAndPreview();
            }
        }

        public string RotationText
        {
            get { return _rotationText; }
            set
            {
                if (_rotationText == value) return;
                _rotationText = value;
                OnPropertyChanged();
                RefreshGeneratedPropertiesAndPreview();
            }
        }

        public bool AutoSpacing
        {
            get { return _autoSpacing; }
            set
            {
                if (_autoSpacing == value) return;
                _autoSpacing = value;
                OnPropertyChanged();
                OnPropertyChanged("IsAutomaticSpacing");
                OnPropertyChanged("SpacingFieldsEnabled");
                UpdateAutomaticSpacing();
                RefreshGeneratedPropertiesAndPreview();
            }
        }

        public bool IsAutomaticSpacing
        {
            get { return AutoSpacing; }
            set { AutoSpacing = value; }
        }

        public bool SpacingFieldsEnabled
        {
            get { return !AutoSpacing; }
        }

        public string PileSpacingText
        {
            get { return _pileSpacingText; }
            set
            {
                if (_pileSpacingText == value) return;
                _pileSpacingText = value;
                OnPropertyChanged();
                RefreshGeneratedPropertiesAndPreview();
            }
        }

        public string SpacingXText
        {
            get { return _spacingXText; }
            set
            {
                if (_spacingXText == value) return;
                _spacingXText = value;
                OnPropertyChanged();
                if (LockSpacingXAndY && !AutoSpacing)
                {
                    _spacingYText = value;
                    OnPropertyChanged("SpacingYText");
                }

                RefreshGeneratedPropertiesAndPreview();
            }
        }

        public string SpacingYText
        {
            get { return _spacingYText; }
            set
            {
                if (_spacingYText == value) return;
                _spacingYText = value;
                OnPropertyChanged();
                RefreshGeneratedPropertiesAndPreview();
            }
        }

        public bool LockSpacingXAndY
        {
            get { return _lockSpacingXAndY; }
            set
            {
                if (_lockSpacingXAndY == value) return;
                _lockSpacingXAndY = value;
                OnPropertyChanged();
                if (value)
                {
                    SpacingYText = SpacingXText;
                }
            }
        }

        public string PileCapThicknessText
        {
            get { return _pileCapThicknessText; }
            set
            {
                if (_pileCapThicknessText == value) return;
                _pileCapThicknessText = value;
                OnPropertyChanged();
                RefreshGeneratedPropertiesAndPreview();
            }
        }

        public string EdgeDistanceText
        {
            get { return _edgeDistanceText; }
            set
            {
                if (_edgeDistanceText == value) return;
                _edgeDistanceText = value;
                OnPropertyChanged();
                RefreshGeneratedPropertiesAndPreview();
            }
        }

        public string SelectedPileMaterial
        {
            get { return _selectedPileMaterial; }
            set
            {
                if (_selectedPileMaterial == value) return;
                _selectedPileMaterial = value;
                OnPropertyChanged();
                RefreshGeneratedPropertiesAndPreview();
            }
        }

        public string SelectedPileCapMaterial
        {
            get { return _selectedPileCapMaterial; }
            set
            {
                if (_selectedPileCapMaterial == value) return;
                _selectedPileCapMaterial = value;
                OnPropertyChanged();
                RefreshGeneratedPropertiesAndPreview();
            }
        }

        public string PilePropertyName
        {
            get { return _pilePropertyName; }
            private set
            {
                if (_pilePropertyName == value) return;
                _pilePropertyName = value;
                OnPropertyChanged();
            }
        }

        public string PileCapPropertyName
        {
            get { return _pileCapPropertyName; }
            private set
            {
                if (_pileCapPropertyName == value) return;
                _pileCapPropertyName = value;
                OnPropertyChanged();
            }
        }

        public string ValidationMessage
        {
            get { return _validationMessage; }
            private set
            {
                if (_validationMessage == value) return;
                _validationMessage = value;
                OnPropertyChanged();
                OnPropertyChanged("ValidationErrors");
                OnPropertyChanged("ValidationVisibility");
                RaiseAssignCanExecuteChanged();
            }
        }

        public string ValidationErrors
        {
            get { return ValidationMessage; }
        }

        public Visibility ValidationVisibility
        {
            get { return string.IsNullOrWhiteSpace(ValidationMessage) ? Visibility.Collapsed : Visibility.Visible; }
        }

        public string StatusMessage
        {
            get { return _statusMessage; }
            private set
            {
                if (_statusMessage == value) return;
                _statusMessage = value;
                OnPropertyChanged();
            }
        }

        public bool IsProcessing
        {
            get { return _isProcessing; }
            private set
            {
                if (_isProcessing == value) return;
                _isProcessing = value;
                OnPropertyChanged();
                OnPropertyChanged("IsAssigning");
                RaiseAssignCanExecuteChanged();
            }
        }

        public bool IsAssigning
        {
            get { return IsProcessing; }
        }

        public PileCapGeometry PreviewGeometry
        {
            get { return _previewGeometry; }
            private set
            {
                if (ReferenceEquals(_previewGeometry, value)) return;
                _previewGeometry = value;
                OnPropertyChanged();
                PreviewVersion++;
            }
        }

        public double PreviewPileDiameter
        {
            get { return _previewPileDiameter; }
            private set
            {
                if (Math.Abs(_previewPileDiameter - value) < 0.000001) return;
                _previewPileDiameter = value;
                OnPropertyChanged();
            }
        }

        public double PreviewPileCapThickness
        {
            get { return _previewPileCapThickness; }
            private set
            {
                if (Math.Abs(_previewPileCapThickness - value) < 0.000001) return;
                _previewPileCapThickness = value;
                OnPropertyChanged();
            }
        }

        public double PreviewEdgeDistance
        {
            get { return _previewEdgeDistance; }
            private set
            {
                if (Math.Abs(_previewEdgeDistance - value) < 0.000001) return;
                _previewEdgeDistance = value;
                OnPropertyChanged();
            }
        }

        public PileCapGeometry MonoCardGeometry
        {
            get { return _monoCardGeometry; }
            private set
            {
                if (ReferenceEquals(_monoCardGeometry, value)) return;
                _monoCardGeometry = value;
                OnPropertyChanged();
            }
        }

        public PileCapGeometry TwoPileCardGeometry
        {
            get { return _twoPileCardGeometry; }
            private set
            {
                if (ReferenceEquals(_twoPileCardGeometry, value)) return;
                _twoPileCardGeometry = value;
                OnPropertyChanged();
            }
        }

        public PileCapGeometry ThreePileCardGeometry
        {
            get { return _threePileCardGeometry; }
            private set
            {
                if (ReferenceEquals(_threePileCardGeometry, value)) return;
                _threePileCardGeometry = value;
                OnPropertyChanged();
            }
        }

        public PileCapGeometry FourPileCardGeometry
        {
            get { return _fourPileCardGeometry; }
            private set
            {
                if (ReferenceEquals(_fourPileCardGeometry, value)) return;
                _fourPileCardGeometry = value;
                OnPropertyChanged();
            }
        }

        public PileCapAssignmentSummaryDto LastAssignmentSummary
        {
            get { return _lastAssignmentSummary; }
            private set
            {
                if (ReferenceEquals(_lastAssignmentSummary, value)) return;
                _lastAssignmentSummary = value;
                OnPropertyChanged();
            }
        }

        public int PreviewVersion
        {
            get { return _previewVersion; }
            private set
            {
                if (_previewVersion == value) return;
                _previewVersion = value;
                OnPropertyChanged();
            }
        }

        public Visibility MonoSpacingVisibility
        {
            get { return SelectedArrangementType == PileCapArrangementType.Mono ? Visibility.Collapsed : Visibility.Visible; }
        }

        public Visibility SingleSpacingVisibility
        {
            get
            {
                return SelectedArrangementType == PileCapArrangementType.TwoPile ||
                       SelectedArrangementType == PileCapArrangementType.ThreePile
                    ? Visibility.Visible
                    : Visibility.Collapsed;
            }
        }

        public Visibility FourSpacingVisibility
        {
            get { return SelectedArrangementType == PileCapArrangementType.FourPile ? Visibility.Visible : Visibility.Collapsed; }
        }

        public string SpacingDescriptionText
        {
            get
            {
                if (SelectedArrangementType == PileCapArrangementType.Mono)
                {
                    return "Mono pile caps do not use pile spacing.";
                }

                return AutoSpacing ? "Automatic spacing is 3D centre-to-centre." : "Manual spacing is enabled.";
            }
        }

        private void RefreshContext()
        {
            RefreshMaterials();
            RefreshSelection();
        }

        private void RefreshSelectionAndRestoreFocus()
        {
            try
            {
                RefreshSelection();
            }
            finally
            {
                _restoreFocus();
            }
        }

        private void RefreshMaterials()
        {
            OperationResult<IReadOnlyList<string>> result = _connectionService.GetConcreteMaterialNames();
            ConcreteMaterials.Clear();
            if (result.IsSuccess && result.Data != null)
            {
                foreach (string material in result.Data)
                {
                    if (!string.IsNullOrWhiteSpace(material))
                    {
                        ConcreteMaterials.Add(material);
                    }
                }
            }

            if (ConcreteMaterials.Count > 0)
            {
                SelectedPileMaterial = SelectDefaultMaterial(ConcreteMaterials);
                SelectedPileCapMaterial = SelectedPileMaterial;
            }
            else
            {
                SelectedPileMaterial = string.Empty;
                SelectedPileCapMaterial = string.Empty;
                StatusMessage = string.IsNullOrWhiteSpace(result.Message)
                    ? "No concrete materials were found in the current ETABS model."
                    : result.Message;
            }
        }

        private void RefreshSelection()
        {
            OperationResult<CSISapModelConnectionInfoDTO> connectionResult = _connectionService.GetCurrentConnection();
            if (connectionResult.IsSuccess && connectionResult.Data != null)
            {
                ConnectedModelName = string.IsNullOrWhiteSpace(connectionResult.Data.ModelFileName)
                    ? "Unknown model"
                    : connectionResult.Data.ModelFileName;
            }

            OperationResult<CSISapModelPresentUnitSystemDTO> unitResult = _connectionService.GetPresentUnitSystem();
            CurrentUnitsText = unitResult.IsSuccess && unitResult.Data != null
                ? FormatUnitSystem(unitResult.Data)
                : "-";

            OperationResult<IReadOnlyList<CsiSelectedObjectDto>> selectionResult = _connectionService.GetSelectedObjectsFromActiveModel();
            if (!selectionResult.IsSuccess || selectionResult.Data == null)
            {
                SelectedPointCount = 0;
                IgnoredNonPointObjectCount = 0;
                SelectedPointNamesText = "-";
                if (!string.IsNullOrWhiteSpace(selectionResult.Message))
                {
                    StatusMessage = selectionResult.Message;
                }

                return;
            }

            var pointNames = new List<string>();
            int ignored = 0;
            foreach (CsiSelectedObjectDto selectedObject in selectionResult.Data)
            {
                if (selectedObject == null)
                {
                    continue;
                }

                if (string.Equals(selectedObject.ObjectType, "Point", StringComparison.OrdinalIgnoreCase))
                {
                    pointNames.Add(selectedObject.UniqueName);
                }
                else
                {
                    ignored++;
                }
            }

            SelectedPointCount = pointNames.Count;
            IgnoredNonPointObjectCount = ignored;
            SelectedPointNamesText = pointNames.Count == 0 ? "-" : string.Join(", ", pointNames);
            StatusMessage = ignored > 0
                ? "Selection refreshed. Non-point objects will be ignored during assignment."
                : "Selection refreshed.";
        }

        private void Assign()
        {
            if (IsProcessing)
            {
                return;
            }

            PileCapAssignmentRequestDto request;
            string validationMessage;
            if (!TryCreateRequest(out request, out validationMessage))
            {
                ValidationMessage = validationMessage;
                StatusMessage = validationMessage;
                return;
            }

            try
            {
                IsProcessing = true;
                StatusMessage = "Assigning pile caps and piles...";
                RefreshSelection();
                if (SelectedPointCount == 0)
                {
                    string noPointMessage = "No ETABS point objects are currently selected. Select one or more point objects in ETABS, then click Assign again.";
                    ValidationMessage = noPointMessage;
                    StatusMessage = noPointMessage;
                    LastAssignmentSummary = null;
                    return;
                }

                OperationResult<PileCapAssignmentSummaryDto> result = _connectionService.QuickCreatePileCaps(request);
                if (!result.IsSuccess || result.Data == null)
                {
                    StatusMessage = string.IsNullOrWhiteSpace(result.Message)
                        ? "Pile-cap assignment failed."
                        : result.Message;
                    LastAssignmentSummary = result.Data;
                    return;
                }

                LastAssignmentSummary = result.Data;
                StatusMessage = FormatAssignmentSummary(result.Data, result.Message);
                ValidationMessage = string.Empty;
            }
            finally
            {
                IsProcessing = false;
                _restoreFocus();
            }
        }

        private bool CanAssign()
        {
            return !IsProcessing && _hasValidInputs;
        }

        private bool TryCreateRequest(out PileCapAssignmentRequestDto request, out string message)
        {
            request = null;
            PileCapInputParameters input;
            if (!TryCreateInputParameters(out input, out message))
            {
                return false;
            }

            IReadOnlyList<string> validationMessages = _validator.Validate(input);
            if (validationMessages.Count > 0)
            {
                message = string.Join(Environment.NewLine, validationMessages);
                return false;
            }

            request = new PileCapAssignmentRequestDto
            {
                ArrangementType = input.ArrangementType,
                PileDiameterMillimeters = input.PileDiameterMillimeters,
                PileLengthMillimeters = input.PileLengthMillimeters,
                PileMaterial = input.PileMaterial,
                RotationDegrees = input.RotationDegrees,
                AutoSpacing = input.AutoSpacing,
                PileSpacingMillimeters = input.PileSpacingMillimeters,
                SpacingXMillimeters = input.SpacingXMillimeters,
                SpacingYMillimeters = input.SpacingYMillimeters,
                PileCapThicknessMillimeters = input.PileCapThicknessMillimeters,
                EdgeDistanceMillimeters = input.EdgeDistanceMillimeters,
                PileCapMaterial = input.PileCapMaterial,
                SelectCreatedObjects = true
            };
            message = string.Empty;
            return true;
        }

        private bool TryCreateInputParameters(out PileCapInputParameters input, out string message)
        {
            input = new PileCapInputParameters();
            message = string.Empty;
            double pileDiameter;
            double pileLength;
            double rotation;
            double pileSpacing = 0;
            double spacingX = 0;
            double spacingY = 0;
            double thickness;
            double edgeDistance;

            if (!TryParsePositiveOrZero(PileDiameterText, out pileDiameter) ||
                !TryParsePositiveOrZero(PileLengthText, out pileLength) ||
                !TryParseNumber(RotationText, out rotation) ||
                !TryParsePositiveOrZero(PileCapThicknessText, out thickness) ||
                !TryParsePositiveOrZero(EdgeDistanceText, out edgeDistance))
            {
                message = "Enter numeric values for all pile and pile-cap dimensions.";
                return false;
            }

            if ((SelectedArrangementType == PileCapArrangementType.TwoPile ||
                 SelectedArrangementType == PileCapArrangementType.ThreePile) &&
                !TryParsePositiveOrZero(PileSpacingText, out pileSpacing))
            {
                message = "Enter a numeric pile spacing.";
                return false;
            }

            if (SelectedArrangementType == PileCapArrangementType.FourPile &&
                (!TryParsePositiveOrZero(SpacingXText, out spacingX) ||
                 !TryParsePositiveOrZero(SpacingYText, out spacingY)))
            {
                message = "Enter numeric Spacing X and Spacing Y values.";
                return false;
            }

            if (AutoSpacing)
            {
                double defaultSpacing = pileDiameter * 3.0;
                pileSpacing = defaultSpacing;
                spacingX = defaultSpacing;
                spacingY = defaultSpacing;
            }

            input.ArrangementType = SelectedArrangementType;
            input.PileDiameterMillimeters = pileDiameter;
            input.PileLengthMillimeters = pileLength;
            input.RotationDegrees = rotation;
            input.AutoSpacing = AutoSpacing;
            input.PileSpacingMillimeters = pileSpacing;
            input.SpacingXMillimeters = spacingX;
            input.SpacingYMillimeters = spacingY;
            input.PileCapThicknessMillimeters = thickness;
            input.EdgeDistanceMillimeters = edgeDistance;
            input.PileMaterial = SelectedPileMaterial;
            input.PileCapMaterial = SelectedPileCapMaterial;
            return true;
        }

        private void RefreshGeneratedPropertiesAndPreview()
        {
            double pileDiameter;
            double pileCapThickness;
            if (TryParseNumber(PileDiameterText, out pileDiameter))
            {
                PilePropertyName = PileCapPropertyNameBuilder.BuildPileFrameSectionName(pileDiameter, SelectedPileMaterial);
                PreviewPileDiameter = pileDiameter;
            }
            else
            {
                PilePropertyName = "P_{Diameter}D_{Material}";
                PreviewPileDiameter = 800;
            }

            if (TryParseNumber(PileCapThicknessText, out pileCapThickness))
            {
                PileCapPropertyName = PileCapPropertyNameBuilder.BuildPileCapAreaSectionName(pileCapThickness, SelectedPileCapMaterial);
                PreviewPileCapThickness = pileCapThickness;
            }
            else
            {
                PileCapPropertyName = "PC_{Thickness}_{Material}";
                PreviewPileCapThickness = 1500;
            }

            double edgeDistance;
            PreviewEdgeDistance = TryParseNumber(EdgeDistanceText, out edgeDistance) ? edgeDistance : 150;

            PileCapInputParameters input;
            string message;
            if (TryCreateInputParameters(out input, out message))
            {
                IReadOnlyList<string> validationMessages = _validator.Validate(input);
                SetHasValidInputs(validationMessages.Count == 0);
                ValidationMessage = validationMessages.Count == 0 ? string.Empty : string.Join(Environment.NewLine, validationMessages);
                try
                {
                    PreviewGeometry = _geometryCalculator.Calculate(input);
                    RefreshCardGeometries(input);
                }
                catch (Exception ex)
                {
                    SetHasValidInputs(false);
                    ValidationMessage = ex.Message;
                    PreviewGeometry = null;
                }
            }
            else
            {
                SetHasValidInputs(false);
                ValidationMessage = message;
                PreviewGeometry = null;
                RefreshCardGeometries(CreateFallbackPreviewInput());
            }

            RaiseAssignCanExecuteChanged();
        }

        private void SetHasValidInputs(bool value)
        {
            if (_hasValidInputs == value)
            {
                return;
            }

            _hasValidInputs = value;
            RaiseAssignCanExecuteChanged();
        }

        private void RefreshCardGeometries(PileCapInputParameters sourceInput)
        {
            if (sourceInput == null)
            {
                return;
            }

            MonoCardGeometry = CalculateCardGeometry(sourceInput, PileCapArrangementType.Mono);
            TwoPileCardGeometry = CalculateCardGeometry(sourceInput, PileCapArrangementType.TwoPile);
            ThreePileCardGeometry = CalculateCardGeometry(sourceInput, PileCapArrangementType.ThreePile);
            FourPileCardGeometry = CalculateCardGeometry(sourceInput, PileCapArrangementType.FourPile);
        }

        private PileCapGeometry CalculateCardGeometry(PileCapInputParameters sourceInput, PileCapArrangementType arrangementType)
        {
            var input = new PileCapInputParameters
            {
                ArrangementType = arrangementType,
                PileDiameterMillimeters = sourceInput.PileDiameterMillimeters,
                PileLengthMillimeters = sourceInput.PileLengthMillimeters,
                RotationDegrees = 0,
                AutoSpacing = sourceInput.AutoSpacing,
                PileSpacingMillimeters = sourceInput.PileSpacingMillimeters,
                SpacingXMillimeters = sourceInput.SpacingXMillimeters,
                SpacingYMillimeters = sourceInput.SpacingYMillimeters,
                PileCapThicknessMillimeters = sourceInput.PileCapThicknessMillimeters,
                EdgeDistanceMillimeters = sourceInput.EdgeDistanceMillimeters,
                PileMaterial = sourceInput.PileMaterial,
                PileCapMaterial = sourceInput.PileCapMaterial
            };

            if (input.PileSpacingMillimeters <= 0)
            {
                input.PileSpacingMillimeters = input.PileDiameterMillimeters * 3.0;
            }

            if (input.SpacingXMillimeters <= 0)
            {
                input.SpacingXMillimeters = input.PileDiameterMillimeters * 3.0;
            }

            if (input.SpacingYMillimeters <= 0)
            {
                input.SpacingYMillimeters = input.SpacingXMillimeters;
            }

            return _geometryCalculator.Calculate(input);
        }

        private PileCapInputParameters CreateFallbackPreviewInput()
        {
            return new PileCapInputParameters
            {
                ArrangementType = SelectedArrangementType,
                PileDiameterMillimeters = PreviewPileDiameter <= 0 ? 800 : PreviewPileDiameter,
                PileLengthMillimeters = 30000,
                RotationDegrees = 0,
                AutoSpacing = true,
                PileSpacingMillimeters = 2400,
                SpacingXMillimeters = 2400,
                SpacingYMillimeters = 2400,
                PileCapThicknessMillimeters = PreviewPileCapThickness <= 0 ? 1500 : PreviewPileCapThickness,
                EdgeDistanceMillimeters = PreviewEdgeDistance < 0 ? 150 : PreviewEdgeDistance,
                PileMaterial = SelectedPileMaterial,
                PileCapMaterial = SelectedPileCapMaterial
            };
        }

        private void UpdateAutomaticSpacing()
        {
            if (!AutoSpacing)
            {
                OnPropertyChanged("SpacingDescriptionText");
                return;
            }

            double pileDiameter;
            if (!TryParseNumber(PileDiameterText, out pileDiameter))
            {
                return;
            }

            string defaultSpacing = (pileDiameter * 3.0).ToString("0.###", CultureInfo.InvariantCulture);
            _pileSpacingText = defaultSpacing;
            _spacingXText = defaultSpacing;
            _spacingYText = defaultSpacing;
            OnPropertyChanged("PileSpacingText");
            OnPropertyChanged("SpacingXText");
            OnPropertyChanged("SpacingYText");
            OnPropertyChanged("SpacingDescriptionText");
        }

        private static string SelectDefaultMaterial(IReadOnlyList<string> materials)
        {
            foreach (string material in materials)
            {
                if (string.Equals(material, "C32/40", StringComparison.Ordinal))
                {
                    return material;
                }
            }

            return materials.Count > 0 ? materials[0] : string.Empty;
        }

        private static bool TryParsePositiveOrZero(string text, out double value)
        {
            return TryParseNumber(text, out value) && value >= 0;
        }

        private static bool TryParseNumber(string text, out double value)
        {
            return double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out value) ||
                   double.TryParse(text, NumberStyles.Float, CultureInfo.CurrentCulture, out value);
        }

        private static string FormatUnitSystem(CSISapModelPresentUnitSystemDTO units)
        {
            return FormatForce(units.ForceUnit) + "-" + FormatLength(units.LengthUnit) + "-" + FormatTemperature(units.TemperatureUnit);
        }

        private static string FormatForce(int forceUnit)
        {
            switch (forceUnit)
            {
                case 1: return "lb";
                case 2: return "kip";
                case 3: return "N";
                case 4: return "kN";
                case 5: return "kgf";
                case 6: return "tonf";
                default: return "?";
            }
        }

        private static string FormatLength(int lengthUnit)
        {
            switch (lengthUnit)
            {
                case 1: return "inch";
                case 2: return "ft";
                case 3: return "micron";
                case 4: return "mm";
                case 5: return "cm";
                case 6: return "m";
                default: return "?";
            }
        }

        private static string FormatTemperature(int temperatureUnit)
        {
            switch (temperatureUnit)
            {
                case 1: return "F";
                case 2: return "C";
                default: return "?";
            }
        }

        private static string FormatAssignmentSummary(PileCapAssignmentSummaryDto summary, string message)
        {
            var builder = new StringBuilder();
            if (!string.IsNullOrWhiteSpace(message))
            {
                builder.AppendLine(message);
            }

            builder.AppendLine("Pile property: " + summary.PilePropertyName);
            builder.AppendLine("Pile-cap property: " + summary.PileCapPropertyName);

            if (summary.Warnings.Count > 0)
            {
                builder.AppendLine("Warnings:");
                foreach (string warning in summary.Warnings)
                {
                    builder.AppendLine("- " + warning);
                }
            }

            if (summary.Errors.Count > 0)
            {
                builder.AppendLine("Errors:");
                foreach (string error in summary.Errors)
                {
                    builder.AppendLine("- " + error);
                }
            }

            return builder.ToString().Trim();
        }

        private void RaiseAssignCanExecuteChanged()
        {
            IRelayCommand relay = AssignCommand as IRelayCommand;
            if (relay != null)
            {
                relay.RaiseCanExecuteChanged();
            }
        }

        private void SelectArrangement(object parameter)
        {
            if (parameter is PileCapArrangementType)
            {
                SelectedArrangement = (PileCapArrangementType)parameter;
                return;
            }

            string text = parameter as string;
            if (!string.IsNullOrWhiteSpace(text))
            {
                PileCapArrangementType arrangement;
                if (Enum.TryParse(text, true, out arrangement))
                {
                    SelectedArrangement = arrangement;
                }
            }
        }
    }
}
