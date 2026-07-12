using System;
using System.ComponentModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Globalization;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using ExcelCSIToolBox.Application.Interfaces.Etabs;
using ExcelCSIToolBox.Application.Interfaces.Etabs.MiscellaneousData;
using ExcelCSIToolBoxAddIn.UI.Common.Commands;
using ExcelCSIToolBox.Core.Common.Results;
using ExcelCSIToolBox.Core.Models.AnalysisResults;
using ExcelCSIToolBox.Core.Models.ElementConnectivity;
using ExcelCSIToolBox.Core.Models.MiscellaneousData;
using ExcelCSIToolBox.Application.Composition;
using ExcelCSIToolBox.Application.Features.Connectivity;
using ExcelCSIToolBox.Core.Abstractions.CSI;
using ExcelCSIToolBox.Core.Abstractions.Excel;
using ExcelCSIToolBox.Core.Contracts.CSI;
using ExcelCSIToolBox.Infrastructure.CSISapModel;
using ExcelCSIToolBox.Infrastructure.Excel;
using ExcelCSIToolBox.Infrastructure.Services.Etabs.AnalysisResults;
using ExcelCSIToolBox.Infrastructure.Services.Etabs.ElementConnectivity;
using ExcelCSIToolBox.Infrastructure.Services.Etabs.MiscellaneousData;
using ExcelCSIToolBoxAddIn.AddIn.Composition;

namespace ExcelCSIToolBoxAddIn.UI.ViewModels
{
    /// <summary>
    /// ViewModel for CSI toolbox shells.
    /// Exposes connection state, model name, and point/frame placeholder commands.
    /// </summary>
    public partial class CsiToolboxViewModel : ViewModelBase
    {
        private readonly CsiToolboxUseCaseBundle _useCases;
        private readonly ICSISapModelConnectionService _csiConnectionService;
        private readonly IExcelSelectionService _excelSelectionService;
        private readonly IExcelOutputService _excelOutputService;
        private readonly IEtabsMiscellaneousDataRouter _miscellaneousDataRouter;
        private readonly ExportSelectedObjectConnectivityUseCase _exportSelectedObjectConnectivity;
        private readonly IEtabsUnitService _etabsUnitService;

        private string _modelName;
        private bool _isConnected;
        private string _statusText;
        private string _currentModelUnitText;
        private string _modelPath;
        private int _activeWorkspacePage;
        private readonly string _productName;
        private EtabsUnitSystem _selectedUnitSystem;
        private CsiRunningInstanceViewModel _selectedRunningCsiInstance;
        private bool _isInitializingUnitSystems;
        private bool _isApplyingGlobalUnit;
        private bool _isModelLocked;
        private bool _isRefreshingRunningCsiInstances;
        private bool _hasRunningCsiInstance;

        public CsiToolboxViewModel(
            ICSISapModelConnectionService csiConnectionService,
            IExcelSelectionService excelSelectionService,
            IExcelOutputService excelOutputService)
            : this(
                new CsiToolboxUseCaseBundle(csiConnectionService, excelSelectionService, excelOutputService),
                csiConnectionService,
                excelSelectionService,
                excelOutputService)
        {
        }

        public CsiToolboxViewModel(
            CsiToolboxUseCaseBundle useCases,
            ICSISapModelConnectionService csiConnectionService,
            IExcelSelectionService excelSelectionService,
            IExcelOutputService excelOutputService,
            EtabsAnalysisResultServices analysisResultServices = null,
            EtabsElementConnectivityServices elementConnectivityServices = null,
            EtabsMiscellaneousDataServices miscellaneousDataServices = null)
        {
            if (csiConnectionService == null) throw new ArgumentNullException(nameof(csiConnectionService));

            _productName = string.IsNullOrWhiteSpace(csiConnectionService.ProductName)
                ? "CSI"
                : csiConnectionService.ProductName;
            _useCases = useCases ?? throw new ArgumentNullException(nameof(useCases));
            _csiConnectionService = csiConnectionService;
            _excelSelectionService = excelSelectionService ?? throw new ArgumentNullException(nameof(excelSelectionService));
            _excelOutputService = excelOutputService ?? throw new ArgumentNullException(nameof(excelOutputService));
            analysisResultServices = analysisResultServices ?? AppServiceFactory.CreateAnalysisResultServices(csiConnectionService);
            IEtabsUnitService sharedEtabsUnitService = analysisResultServices.UnitService;
            elementConnectivityServices = elementConnectivityServices ?? AppServiceFactory.CreateElementConnectivityServices(csiConnectionService, sharedEtabsUnitService);
            miscellaneousDataServices = miscellaneousDataServices ?? AppServiceFactory.CreateMiscellaneousDataServices(csiConnectionService, sharedEtabsUnitService);
            _miscellaneousDataRouter = miscellaneousDataServices.Router;
            _exportSelectedObjectConnectivity = elementConnectivityServices.ExportSelectedObjectConnectivity;
            _etabsUnitService = sharedEtabsUnitService;
            AnalysisResults = new AnalysisResultsViewModel(
                () => CanUseActiveModel,
                CanExecuteEtabsAction,
                RunAnalysisResult,
                RunEtabsTableItem,
                OpenGetBaseReactionsDialog,
                OpenModalMassParticipationRatiosDialog,
                OpenStoryForcesDialog,
                OpenStoryDriftsDialog,
                OpenStoryMaxOverAverageDisplacementsDialog,
                OpenStoryMaxOverAverageDriftsDialog,
                OpenMassSummaryByStoryDialog);
            AnalysisResults.PropertyChanged += OnAnalysisResultsPropertyChanged;

            LoadCombinations = new System.Collections.ObjectModel.ObservableCollection<ExcelCSIToolBox.Core.Contracts.CSI.CSISapModelLoadCombinationDTO>();
            LoadPatterns = new System.Collections.ObjectModel.ObservableCollection<ExcelCSIToolBox.Core.Contracts.CSI.CSISapModelLoadPatternDTO>();
            FrameSections = new System.Collections.ObjectModel.ObservableCollection<CSISapModelFrameSectionDTO>();
            SectionDimensionAnnotations = new System.Collections.ObjectModel.ObservableCollection<SectionDimensionAnnotation>();
            RunningCsiInstances = new ObservableCollection<CsiRunningInstanceViewModel>();
            InitializeStiffnessModifierPage();
            InitializeModellingHelperPage();
            AvailableUnitSystems = CreateAvailableUnitSystems();
            _isInitializingUnitSystems = true;
            SelectedUnitSystem = AvailableUnitSystems[0];
            _isInitializingUnitSystems = false;

            RefreshRunningCsiInstancesCommand = new RelayCommand(RefreshRunningCsiInstancesFromUi);
            AttachToRunningCsiCommand = new RelayCommand(() => LoadConnectionState(showMessage: true), () => HasRunningCsiInstance);
            CloseCurrentInstanceCommand = new RelayCommand(CloseCurrentInstance, CanExecuteCsiAction);
            ToggleModelLockCommand = new RelayCommand(ToggleModelLock, CanExecuteCsiAction);
            SelectWorkspacePageCommand = new RelayCommand<string>(SelectWorkspacePage);
            ExportAnalysisResultTableCommand = AnalysisResults.ExportAnalysisResultTableCommand;
            ExportEtabsTableItemCommand = AnalysisResults.ExportEtabsTableItemCommand;
            RefreshFrameStiffnessSectionsCommand = new RelayCommand(RefreshFrameStiffnessSections, CanExecuteCsiAction);
            RefreshAreaStiffnessSectionsCommand = new RelayCommand(RefreshAreaStiffnessSections, CanExecuteCsiAction);
            SelectVisibleFrameStiffnessSectionsCommand = new RelayCommand(SelectVisibleFrameStiffnessSections, CanExecuteCsiAction);
            ClearFrameStiffnessSectionSelectionCommand = new RelayCommand(ClearFrameStiffnessSectionSelection, CanExecuteCsiAction);
            SelectVisibleAreaStiffnessSectionsCommand = new RelayCommand(SelectVisibleAreaStiffnessSections, CanExecuteCsiAction);
            ClearAreaStiffnessSectionSelectionCommand = new RelayCommand(ClearAreaStiffnessSectionSelection, CanExecuteCsiAction);
            ApplyFrameStiffnessModifiersCommand = new RelayCommand(ApplyFrameStiffnessModifiers, CanExecuteCsiAction);
            ApplyAreaStiffnessModifiersCommand = new RelayCommand(ApplyAreaStiffnessModifiers, CanExecuteCsiAction);
            ResetFrameStiffnessModifiersCommand = new RelayCommand(ResetFrameModifierFields, CanExecuteCsiAction);
            ResetAreaStiffnessModifiersCommand = new RelayCommand(ResetAreaModifierFields, CanExecuteCsiAction);
            OpenCreateSectionDialogCommand = new RelayCommand(OpenCreateSectionDialog, CanExecuteCsiAction);

            CreateIshapeSectionCommand = new RelayCommand(() => ShowOperationResult(_useCases.CreateSteelISections.Execute()), CanExecuteCsiAction);
            CreateChannelSectionCommand = new RelayCommand(() => ShowOperationResult(_useCases.CreateSteelChannelSections.Execute()), CanExecuteCsiAction);
            CreateAngleSectionCommand = new RelayCommand(() => ShowOperationResult(_useCases.CreateSteelAngleSections.Execute()), CanExecuteCsiAction);
            CreateTubeSectionCommand = new RelayCommand(() => ShowOperationResult(_useCases.CreateSteelTubeSections.Execute()), CanExecuteCsiAction);
            CreatePipeSectionCommand = new RelayCommand(() => ShowOperationResult(_useCases.CreateSteelPipeSections.Execute()), CanExecuteCsiAction);

            CreateConcreteRectangleSectionCommand = new RelayCommand(() => ShowOperationResult(_useCases.CreateConcreteRectangleSections.Execute()), CanExecuteCsiAction);
            CreateConcreteCircleSectionCommand = new RelayCommand(() => ShowOperationResult(_useCases.CreateConcreteCircleSections.Execute()), CanExecuteCsiAction);

            SelectPointsByUniqueNameCommand = new RelayCommand(SelectPointsByUniqueName, CanExecuteCsiAction);
            SelectFramesByUniqueNameCommand = new RelayCommand(SelectFramesByUniqueName, CanExecuteCsiAction);
            AddPointByCartesianCommand = new RelayCommand(AddPointByCartesian, CanExecuteCsiAction);
            SetPointsCommand = new RelayCommand(() => ShowPlaceholder("Set Points"), CanExecuteCsiAction);
            RenameSelectedPointsCommand = new RelayCommand(() => ShowPlaceholder("Rename Selected Points"), CanExecuteCsiAction);
            GetSelectedPointsCommand = new RelayCommand(GetSelectedPoints, CanExecuteCsiAction);

            AddFramesByCoordinatesCommand = new RelayCommand(AddFramesByCoordinates, CanExecuteCsiAction);
            AddFramesByPointNamesCommand = new RelayCommand(AddFramesByPointNames, CanExecuteCsiAction);
            SetFramesCommand = new RelayCommand(() => ShowPlaceholder("Set Frames"), CanExecuteCsiAction);
            RenameFramesCommand = new RelayCommand(() => ShowPlaceholder("Rename Frames"), CanExecuteCsiAction);
            GetSelectedFramesCommand = new RelayCommand(GetSelectedFrames, CanExecuteCsiAction);
            GetFrameSectionPropertyCommand = new RelayCommand(() => ShowPlaceholder("Get Frame Section Property"), CanExecuteCsiAction);
            SetFrameSectionPropertyCommand = new RelayCommand(() => ShowPlaceholder("Set Frame Section Property"), CanExecuteCsiAction);
            GetFrameGroupAssignmentCommand = new RelayCommand(() => ShowPlaceholder("Get Frame Group Assignment"), CanExecuteCsiAction);
            SetFrameGroupAssignmentCommand = new RelayCommand(() => ShowPlaceholder("Set Frame Group Assignment"), CanExecuteCsiAction);
            GetFrameModifierCommand = new RelayCommand(() => ShowPlaceholder("Get Frame Modifier"), CanExecuteCsiAction);
            SetFrameModifierCommand = new RelayCommand(() => ShowPlaceholder("Set Frame Modifier"), CanExecuteCsiAction);
            CreateShellAreasFromSelectedFramesCommand = new RelayCommand(CreateShellAreasFromSelectedFrames, CanExecuteCsiAction);
            GetPointGroupAssignmentCommand = new RelayCommand(() => ShowPlaceholder("Get Point Group Assignment"), CanExecuteCsiAction);
            SetPointGroupAssignmentCommand = new RelayCommand(() => ShowPlaceholder("Set Point Group Assignment"), CanExecuteCsiAction);

            GetLoadPatternsCommand = new RelayCommand(GetLoadPatterns, CanExecuteCsiAction);
            AddLoadPatternFromExcelCommand = new RelayCommand(() => ShowPlaceholder("Add Load Pattern From Excel"), CanExecuteCsiAction);
            DeleteSelectedLoadPatternsCommand = new RelayCommand<System.Collections.IList>(DeleteSelectedLoadPatterns, _ => CanUseActiveModel);
            
            GetLoadCombinationsCommand = new RelayCommand(GetLoadCombinations, CanExecuteCsiAction);
            ModifyLoadCombinationsInMatrixViewCommand = new RelayCommand(ModifyLoadCombinationsInMatrixView, CanExecuteCsiAction);
            ExportLoadCombinationMatrixToExcelCommand = new RelayCommand(ExportLoadCombinationMatrixToExcel, CanExecuteCsiAction);
            OpenShellUniformLoadSetFormCommand = new RelayCommand(OpenShellUniformLoadSetForm, CanExecuteEtabsAction);
            ExportShellUniformLoadSetDefinitionsCommand = new RelayCommand(ExportShellUniformLoadSetDefinitions, CanExecuteEtabsAction);
            AddLoadCombinationFromExcelCommand = ModifyLoadCombinationsInMatrixViewCommand;
            DeleteSelectedLoadCombinationsCommand = new RelayCommand<System.Collections.IList>(DeleteSelectedLoadCombinations, _ => CanUseActiveModel);
            ViewLoadCombinationCommand = new RelayCommand<System.Collections.IList>(ViewLoadCombination, _ => CanUseActiveModel);

            GetBaseReactionsCommand = AnalysisResults.GetBaseReactionsCommand;
            GetModalMassParticipationRatiosCommand = AnalysisResults.GetModalMassParticipationRatiosCommand;
            GetStoryForcesCommand = AnalysisResults.GetStoryForcesCommand;
            GetStoryDriftsCommand = AnalysisResults.GetStoryDriftsCommand;
            GetStoryMaxOverAverageDisplacementsCommand = AnalysisResults.GetStoryMaxOverAverageDisplacementsCommand;
            GetStoryMaxOverAverageDriftsCommand = AnalysisResults.GetStoryMaxOverAverageDriftsCommand;
            GetMassSummaryByStoryCommand = AnalysisResults.GetMassSummaryByStoryCommand;
            
            GetFrameSectionsCommand = new RelayCommand(GetFrameSections, CanExecuteCsiAction);
            EditFrameSectionCommand = new RelayCommand<CSISapModelFrameSectionDTO>(EditFrameSection, _ => CanUseActiveModel);

            CurrentModelUnitText = "Not yet attached";
            SetTableGroup("ANALYSIS RESULTS", "Base Reactions");
            LoadConnectionState(showMessage: false);
        }

        private void OnAnalysisResultsPropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (e == null || string.IsNullOrWhiteSpace(e.PropertyName))
            {
                return;
            }

            OnPropertyChanged(e.PropertyName);
            if (e.PropertyName == nameof(AnalysisResultsViewModel.ActiveTableCategory))
            {
                OnPropertyChanged(nameof(ActivePageBreadcrumb));
            }
            else if (e.PropertyName == nameof(AnalysisResultsViewModel.ActiveAnalysisResultsGroup))
            {
                OnPropertyChanged(nameof(ActivePageTitle));
                OnPropertyChanged(nameof(ActivePageBreadcrumb));
            }
            else if (e.PropertyName == nameof(AnalysisResultsViewModel.SelectedAnalysisResultTable))
            {
                OnPropertyChanged(nameof(AnalysisResultPlaceholderText));
            }
        }

        public string ModelName
        {
            get { return _modelName; }
            private set
            {
                _modelName = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(ModelDisplayText));
            }
        }

        public bool IsConnected
        {
            get { return _isConnected; }
            private set
            {
                _isConnected = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(CanUseActiveModel));
                RefreshCommandStates();
            }
        }

        public bool HasRunningCsiInstance
        {
            get { return _hasRunningCsiInstance; }
            private set
            {
                if (_hasRunningCsiInstance == value)
                {
                    return;
                }

                _hasRunningCsiInstance = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(CanUseActiveModel));
                RefreshCommandStates();
            }
        }

        public bool CanUseActiveModel
        {
            get { return IsConnected && HasRunningCsiInstance; }
        }

        private bool CanExecuteCsiAction()
        {
            return CanUseActiveModel;
        }

        private bool CanExecuteEtabsAction()
        {
            return IsEtabs && CanUseActiveModel;
        }

        public bool IsModelLocked
        {
            get { return _isModelLocked; }
            private set
            {
                if (_isModelLocked == value)
                {
                    return;
                }

                _isModelLocked = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(ModelLockButtonText));
            }
        }

        public string ModelLockButtonText
        {
            get { return IsModelLocked ? "Unlock Model" : "Lock Model"; }
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

        public string CurrentModelUnitText
        {
            get
            {
                if (IsSap2000)
                {
                    return "kN-m-C";
                }
                return string.IsNullOrWhiteSpace(_currentModelUnitText) ? "Not yet attached" : _currentModelUnitText;
            }
            private set
            {
                _currentModelUnitText = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(OffsetLengthUnitText));
                RefreshSectionDimensionAnnotations();
            }
        }

        public System.Collections.ObjectModel.ObservableCollection<EtabsUnitSystem> AvailableUnitSystems { get; private set; }

        public EtabsUnitSystem SelectedUnitSystem
        {
            get { return _selectedUnitSystem; }
            set
            {
                if (ReferenceEquals(_selectedUnitSystem, value))
                {
                    SyncSelectedUnitService(value);
                    return;
                }

                _selectedUnitSystem = value;
                OnPropertyChanged();
                if (value != null)
                {
                    CurrentModelUnitText = value.PresentUnitsText;
                }

                SyncSelectedUnitService(value);

                if (!_isInitializingUnitSystems && IsConnected)
                {
                    ApplySelectedGlobalUnit(showMessages: true);
                }
            }
        }

        private void SyncSelectedUnitService(EtabsUnitSystem unitSystem)
        {
            _etabsUnitService.SetSelectedUnitSystem(
                unitSystem == null ? null : unitSystem.ToDto(),
                unitSystem == null ? null : unitSystem.DisplayName,
                unitSystem == null ? null : unitSystem.PresentUnitsText);
        }

        public string ModelPath
        {
            get { return _modelPath; }
            private set
            {
                _modelPath = value;
                OnPropertyChanged();
            }
        }

        public string ModelDisplayText => $"{ModelName}";

        public CsiProductType ProductType => string.Equals(_productName, "SAP2000", StringComparison.OrdinalIgnoreCase) 
            ? CsiProductType.SAP2000 
            : CsiProductType.ETABS;

        public bool IsSap2000 => ProductType == CsiProductType.SAP2000;

        public bool IsEtabs => ProductType == CsiProductType.ETABS;

        public string ProductTitle => $"{_productName} Toolbox";

        public string ObjectConnectivityTitle => $"{_productName} Object Connectivity";

        public Visibility EtabsVisibility => IsEtabs ? Visibility.Visible : Visibility.Collapsed;

        public Visibility Sap2000Visibility => IsSap2000 ? Visibility.Visible : Visibility.Collapsed;

        public ObservableCollection<CsiRunningInstanceViewModel> RunningCsiInstances { get; private set; }

        public CsiRunningInstanceViewModel SelectedRunningCsiInstance
        {
            get { return _selectedRunningCsiInstance; }
            set
            {
                if (ReferenceEquals(_selectedRunningCsiInstance, value))
                {
                    return;
                }

                _selectedRunningCsiInstance = value;
                OnPropertyChanged();

                if (!_isRefreshingRunningCsiInstances && value != null)
                {
                    AttachToSelectedRunningInstance(value, true);
                }
            }
        }

        public int ActiveWorkspacePage
        {
            get { return _activeWorkspacePage; }
            set
            {
                if (_activeWorkspacePage == value)
                {
                    return;
                }

                _activeWorkspacePage = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(ActivePageTitle));
                OnPropertyChanged(nameof(ActivePageBreadcrumb));
            }
        }

        public string ActivePageTitle
        {
            get
            {
                switch (ActiveWorkspacePage)
                {
                    case 1: return "Point Tools";
                    case 2: return "Frame Tools";
                    case 3: return "Shell Tools";
                    case 4: return "Load Pattern";
                    case 5: return "Load Combination";
                    case 6: return string.IsNullOrWhiteSpace(ActiveAnalysisResultsGroup) ? "Analysis Results" : ActiveAnalysisResultsGroup;
                    case 7: return "Section Property - Stiffness Modifier";
                    case 8: return "Helpers";
                    case 9: return "Shell Uniform Load Set Manager";
                    default: return "Section Property";
                }
            }
        }

        public string ActivePageBreadcrumb
        {
            get
            {
                string prefix = ProductTitle;
                if (ActiveWorkspacePage == 7)
                {
                    return $"{prefix} / General Information / Section Property - Stiffness Modifier";
                }

                if (ActiveWorkspacePage == 8)
                {
                    return $"{prefix} / MODELLING HELPER / Helpers";
                }

                if (ActiveWorkspacePage == 9)
                {
                    return $"{prefix} / Model / Shell Uniform Load Set Manager";
                }

                return ActiveWorkspacePage == 6
                    ? $"{prefix} / {ActiveTableCategory} / {ActivePageTitle}"
                    : $"{prefix} / {ActivePageTitle}";
            }
        }

        public string ActiveTableCategory
        {
            get { return AnalysisResults.ActiveTableCategory; }
            private set
            {
                AnalysisResults.ActiveTableCategory = value;
            }
        }

        public string ActiveAnalysisResultsGroup
        {
            get { return AnalysisResults.ActiveAnalysisResultsGroup; }
            private set
            {
                AnalysisResults.ActiveAnalysisResultsGroup = value;
            }
        }

        public string SelectedAnalysisResultTable
        {
            get { return AnalysisResults.SelectedAnalysisResultTable; }
            set
            {
                AnalysisResults.SelectedAnalysisResultTable = value;
            }
        }

        public string AnalysisResultPlaceholderText
        {
            get { return AnalysisResults.AnalysisResultPlaceholderText; }
        }

        public ICommand AttachToRunningCsiCommand { get; }
        public ICommand RefreshRunningCsiInstancesCommand { get; }
        public ICommand CloseCurrentInstanceCommand { get; }
        public ICommand ToggleModelLockCommand { get; }
        public ICommand SelectWorkspacePageCommand { get; }
        public ICommand ExportAnalysisResultTableCommand { get; }
        public ICommand ExportEtabsTableItemCommand { get; }
        public ICommand RefreshFrameStiffnessSectionsCommand { get; }
        public ICommand RefreshAreaStiffnessSectionsCommand { get; }
        public ICommand SelectVisibleFrameStiffnessSectionsCommand { get; }
        public ICommand ClearFrameStiffnessSectionSelectionCommand { get; }
        public ICommand SelectVisibleAreaStiffnessSectionsCommand { get; }
        public ICommand ClearAreaStiffnessSectionSelectionCommand { get; }
        public ICommand ApplyFrameStiffnessModifiersCommand { get; }
        public ICommand ApplyAreaStiffnessModifiersCommand { get; }
        public ICommand ResetFrameStiffnessModifiersCommand { get; }
        public ICommand ResetAreaStiffnessModifiersCommand { get; }

        public ICommand CreateIshapeSectionCommand { get; }
        public ICommand CreateChannelSectionCommand { get; }
        public ICommand CreateAngleSectionCommand { get; }
        public ICommand CreateTubeSectionCommand { get; }
        public ICommand CreatePipeSectionCommand { get; }

        public ICommand CreateConcreteRectangleSectionCommand { get; }
        public ICommand CreateConcreteCircleSectionCommand { get; }

        public ICommand SelectPointsByUniqueNameCommand { get; }
        public ICommand SelectFramesByUniqueNameCommand { get; }
        public ICommand AddPointByCartesianCommand { get; }
        public ICommand SetPointsCommand { get; }
        public ICommand RenameSelectedPointsCommand { get; }
        public ICommand GetSelectedPointsCommand { get; }

        public ICommand AddFramesByCoordinatesCommand { get; }
        public ICommand AddFramesByPointNamesCommand { get; }
        public ICommand SetFramesCommand { get; }
        public ICommand RenameFramesCommand { get; }
        public ICommand GetSelectedFramesCommand { get; }
        public ICommand GetFrameSectionPropertyCommand { get; }
        public ICommand SetFrameSectionPropertyCommand { get; }
        public ICommand GetFrameGroupAssignmentCommand { get; }
        public ICommand SetFrameGroupAssignmentCommand { get; }
        public ICommand GetFrameModifierCommand { get; }
        public ICommand SetFrameModifierCommand { get; }
        public ICommand CreateShellAreasFromSelectedFramesCommand { get; }
        public ICommand GetPointGroupAssignmentCommand { get; }
        public ICommand SetPointGroupAssignmentCommand { get; }

        public ICommand GetLoadPatternsCommand { get; }
        public ICommand AddLoadPatternFromExcelCommand { get; }
        public ICommand DeleteSelectedLoadPatternsCommand { get; }
        
        public ICommand GetLoadCombinationsCommand { get; }
        public ICommand ModifyLoadCombinationsInMatrixViewCommand { get; }
        public ICommand ExportLoadCombinationMatrixToExcelCommand { get; }
        public ICommand OpenShellUniformLoadSetFormCommand { get; }
        public ICommand ExportShellUniformLoadSetDefinitionsCommand { get; }
        public ICommand AddLoadCombinationFromExcelCommand { get; }
        public ICommand DeleteSelectedLoadCombinationsCommand { get; }
        public ICommand ViewLoadCombinationCommand { get; }
        public ICommand GetBaseReactionsCommand { get; }
        public ICommand GetModalMassParticipationRatiosCommand { get; }
        public ICommand GetStoryForcesCommand { get; }
        public ICommand GetStoryDriftsCommand { get; }
        public ICommand GetStoryMaxOverAverageDisplacementsCommand { get; }
        public ICommand GetStoryMaxOverAverageDriftsCommand { get; }
        public ICommand GetMassSummaryByStoryCommand { get; }
        
        public ICommand GetFrameSectionsCommand { get; }
        public ICommand EditFrameSectionCommand { get; }
        public ICommand OpenCreateSectionDialogCommand { get; }

        public System.Collections.ObjectModel.ObservableCollection<ExcelCSIToolBox.Core.Contracts.CSI.CSISapModelLoadPatternDTO> LoadPatterns { get; }
        public System.Collections.ObjectModel.ObservableCollection<ExcelCSIToolBox.Core.Contracts.CSI.CSISapModelLoadCombinationDTO> LoadCombinations { get; }
        public System.Collections.ObjectModel.ObservableCollection<CSISapModelFrameSectionDTO> FrameSections { get; }
        public System.Collections.ObjectModel.ObservableCollection<SectionDimensionAnnotation> SectionDimensionAnnotations { get; }
        public AnalysisResultsViewModel AnalysisResults { get; private set; }
        public System.Collections.ObjectModel.ObservableCollection<AnalysisResultItem> AnalysisResultTables
        {
            get { return AnalysisResults.AnalysisResultTables; }
        }

        public System.Collections.ObjectModel.ObservableCollection<object> EtabsTableItems
        {
            get { return AnalysisResults.EtabsTableItems; }
        }

        private CSISapModelFrameSectionDTO _selectedFrameSection;
        public CSISapModelFrameSectionDTO SelectedFrameSection
        {
            get => _selectedFrameSection;
            set
            {
                _selectedFrameSection = value;
                OnPropertyChanged();
                LoadSelectedSectionDetail(value);
            }
        }

        private CSISapModelFrameSectionDetailDTO _selectedFrameSectionDetail;
        public CSISapModelFrameSectionDetailDTO SelectedFrameSectionDetail
        {
            get => _selectedFrameSectionDetail;
            private set
            {
                _selectedFrameSectionDetail = value;
                OnPropertyChanged();
                RefreshSectionDimensionAnnotations();
            }
        }

    }
}
