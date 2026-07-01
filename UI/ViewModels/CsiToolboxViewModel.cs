using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Globalization;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using ExcelCSIToolBox.Application.Interfaces.Etabs;
using ExcelCSIToolBox.Application.Interfaces.Etabs.AnalysisResults;
using ExcelCSIToolBox.Application.Interfaces.Etabs.ElementConnectivity;
using ExcelCSIToolBox.Application.Interfaces.Etabs.MiscellaneousData;
using ExcelCSIToolBox.Core.Common.Commands;
using ExcelCSIToolBox.Core.Common.Results;
using ExcelCSIToolBox.Core.Models.AnalysisResults;
using ExcelCSIToolBox.Core.Models.ElementConnectivity;
using ExcelCSIToolBox.Core.Models.MiscellaneousData;
using ExcelCSIToolBox.Application.UseCases;
using ExcelCSIToolBox.Core.Abstractions.CSI;
using ExcelCSIToolBox.Core.Abstractions.Excel;
using ExcelCSIToolBox.Data;
using ExcelCSIToolBox.Data.DTOs.CSI;
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
        private readonly IEtabsAnalysisResultRouter _analysisResultRouter;
        private readonly IEtabsElementConnectivityRouter _elementConnectivityRouter;
        private readonly IEtabsMiscellaneousDataRouter _miscellaneousDataRouter;
        private readonly IEtabsUnitService _etabsUnitService;

        private string _modelName;
        private bool _isConnected;
        private string _statusText;
        private string _currentModelUnitText;
        private string _modelPath;
        private int _activeWorkspacePage;
        private string _activeTableCategory;
        private string _activeAnalysisResultsGroup;
        private string _selectedAnalysisResultTable;
        private readonly string _productName;
        private EtabsUnitSystem _selectedUnitSystem;
        private CsiRunningInstanceViewModel _selectedRunningCsiInstance;
        private bool _isInitializingUnitSystems;
        private bool _isApplyingGlobalUnit;
        private bool _isModelLocked;
        private bool _isRefreshingRunningCsiInstances;

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
            elementConnectivityServices = elementConnectivityServices ?? AppServiceFactory.CreateElementConnectivityServices(csiConnectionService);
            miscellaneousDataServices = miscellaneousDataServices ?? AppServiceFactory.CreateMiscellaneousDataServices(csiConnectionService);
            _analysisResultRouter = analysisResultServices.Router;
            _elementConnectivityRouter = elementConnectivityServices.Router;
            _miscellaneousDataRouter = miscellaneousDataServices.Router;
            _etabsUnitService = analysisResultServices.UnitService;

            LoadCombinations = new System.Collections.ObjectModel.ObservableCollection<ExcelCSIToolBox.Data.DTOs.CSI.CSISapModelLoadCombinationDTO>();
            LoadPatterns = new System.Collections.ObjectModel.ObservableCollection<ExcelCSIToolBox.Data.DTOs.CSI.CSISapModelLoadPatternDTO>();
            FrameSections = new System.Collections.ObjectModel.ObservableCollection<CSISapModelFrameSectionDTO>();
            SectionDimensionAnnotations = new System.Collections.ObjectModel.ObservableCollection<SectionDimensionAnnotation>();
            AnalysisResultTables = new System.Collections.ObjectModel.ObservableCollection<AnalysisResultItem>();
            EtabsTableItems = new System.Collections.ObjectModel.ObservableCollection<object>();
            RunningCsiInstances = new ObservableCollection<CsiRunningInstanceViewModel>();
            InitializeStiffnessModifierPage();
            InitializeModellingHelperPage();
            AvailableUnitSystems = CreateAvailableUnitSystems();
            _isInitializingUnitSystems = true;
            SelectedUnitSystem = AvailableUnitSystems[0];
            _isInitializingUnitSystems = false;

            RefreshRunningCsiInstancesCommand = new RelayCommand(RefreshRunningCsiInstancesFromUi);
            AttachToRunningCsiCommand = new RelayCommand(() => LoadConnectionState(showMessage: true));
            CloseCurrentInstanceCommand = new RelayCommand(CloseCurrentInstance, () => IsConnected);
            ToggleModelLockCommand = new RelayCommand(ToggleModelLock, () => IsConnected);
            SelectWorkspacePageCommand = new RelayCommand<string>(SelectWorkspacePage);
            ExportAnalysisResultTableCommand = new RelayCommand<AnalysisResultItem>(RunAnalysisResult, item => item != null);
            ExportEtabsTableItemCommand = new RelayCommand<object>(RunEtabsTableItem, item => item != null);
            RefreshFrameStiffnessSectionsCommand = new RelayCommand(RefreshFrameStiffnessSections, () => IsConnected);
            RefreshAreaStiffnessSectionsCommand = new RelayCommand(RefreshAreaStiffnessSections, () => IsConnected);
            SelectVisibleFrameStiffnessSectionsCommand = new RelayCommand(SelectVisibleFrameStiffnessSections, () => IsConnected);
            ClearFrameStiffnessSectionSelectionCommand = new RelayCommand(ClearFrameStiffnessSectionSelection, () => IsConnected);
            SelectVisibleAreaStiffnessSectionsCommand = new RelayCommand(SelectVisibleAreaStiffnessSections, () => IsConnected);
            ClearAreaStiffnessSectionSelectionCommand = new RelayCommand(ClearAreaStiffnessSectionSelection, () => IsConnected);
            ApplyFrameStiffnessModifiersCommand = new RelayCommand(ApplyFrameStiffnessModifiers, () => IsConnected);
            ApplyAreaStiffnessModifiersCommand = new RelayCommand(ApplyAreaStiffnessModifiers, () => IsConnected);
            ResetFrameStiffnessModifiersCommand = new RelayCommand(ResetFrameModifierFields);
            ResetAreaStiffnessModifiersCommand = new RelayCommand(ResetAreaModifierFields);

            CreateIshapeSectionCommand = new RelayCommand(() => ShowOperationResult(_useCases.CreateSteelISections.Execute()), () => IsConnected);
            CreateChannelSectionCommand = new RelayCommand(() => ShowOperationResult(_useCases.CreateSteelChannelSections.Execute()), () => IsConnected);
            CreateAngleSectionCommand = new RelayCommand(() => ShowOperationResult(_useCases.CreateSteelAngleSections.Execute()), () => IsConnected);
            CreateTubeSectionCommand = new RelayCommand(() => ShowOperationResult(_useCases.CreateSteelTubeSections.Execute()), () => IsConnected);
            CreatePipeSectionCommand = new RelayCommand(() => ShowOperationResult(_useCases.CreateSteelPipeSections.Execute()), () => IsConnected);

            CreateConcreteRectangleSectionCommand = new RelayCommand(() => ShowOperationResult(_useCases.CreateConcreteRectangleSections.Execute()), () => IsConnected);
            CreateConcreteCircleSectionCommand = new RelayCommand(() => ShowOperationResult(_useCases.CreateConcreteCircleSections.Execute()), () => IsConnected);

            SelectPointsByUniqueNameCommand = new RelayCommand(SelectPointsByUniqueName, () => IsConnected);
            SelectFramesByUniqueNameCommand = new RelayCommand(SelectFramesByUniqueName, () => IsConnected);
            AddPointByCartesianCommand = new RelayCommand(AddPointByCartesian, () => IsConnected);
            SetPointsCommand = new RelayCommand(() => ShowPlaceholder("Set Points"), () => IsConnected);
            RenameSelectedPointsCommand = new RelayCommand(() => ShowPlaceholder("Rename Selected Points"), () => IsConnected);
            GetSelectedPointsCommand = new RelayCommand(GetSelectedPoints, () => IsConnected);

            AddFramesByCoordinatesCommand = new RelayCommand(AddFramesByCoordinates, () => IsConnected);
            AddFramesByPointNamesCommand = new RelayCommand(AddFramesByPointNames, () => IsConnected);
            SetFramesCommand = new RelayCommand(() => ShowPlaceholder("Set Frames"), () => IsConnected);
            RenameFramesCommand = new RelayCommand(() => ShowPlaceholder("Rename Frames"), () => IsConnected);
            GetSelectedFramesCommand = new RelayCommand(GetSelectedFrames, () => IsConnected);
            GetFrameSectionPropertyCommand = new RelayCommand(() => ShowPlaceholder("Get Frame Section Property"), () => IsConnected);
            SetFrameSectionPropertyCommand = new RelayCommand(() => ShowPlaceholder("Set Frame Section Property"), () => IsConnected);
            GetFrameGroupAssignmentCommand = new RelayCommand(() => ShowPlaceholder("Get Frame Group Assignment"), () => IsConnected);
            SetFrameGroupAssignmentCommand = new RelayCommand(() => ShowPlaceholder("Set Frame Group Assignment"), () => IsConnected);
            GetFrameModifierCommand = new RelayCommand(() => ShowPlaceholder("Get Frame Modifier"), () => IsConnected);
            SetFrameModifierCommand = new RelayCommand(() => ShowPlaceholder("Set Frame Modifier"), () => IsConnected);
            CreateShellAreasFromSelectedFramesCommand = new RelayCommand(CreateShellAreasFromSelectedFrames, () => IsConnected);
            GetPointGroupAssignmentCommand = new RelayCommand(() => ShowPlaceholder("Get Point Group Assignment"), () => IsConnected);
            SetPointGroupAssignmentCommand = new RelayCommand(() => ShowPlaceholder("Set Point Group Assignment"), () => IsConnected);

            GetLoadPatternsCommand = new RelayCommand(GetLoadPatterns, () => IsConnected);
            AddLoadPatternFromExcelCommand = new RelayCommand(() => ShowPlaceholder("Add Load Pattern From Excel"), () => IsConnected);
            DeleteSelectedLoadPatternsCommand = new RelayCommand<System.Collections.IList>(DeleteSelectedLoadPatterns, _ => IsConnected);
            
            GetLoadCombinationsCommand = new RelayCommand(GetLoadCombinations, () => IsConnected);
            ModifyLoadCombinationsInMatrixViewCommand = new RelayCommand(ModifyLoadCombinationsInMatrixView, () => IsConnected);
            ExportLoadCombinationMatrixToExcelCommand = new RelayCommand(ExportLoadCombinationMatrixToExcel, () => IsConnected);
            AddLoadCombinationFromExcelCommand = ModifyLoadCombinationsInMatrixViewCommand;
            DeleteSelectedLoadCombinationsCommand = new RelayCommand<System.Collections.IList>(DeleteSelectedLoadCombinations, _ => IsConnected);
            ViewLoadCombinationCommand = new RelayCommand<System.Collections.IList>(ViewLoadCombination, _ => IsConnected);

            GetBaseReactionsCommand = new RelayCommand(OpenGetBaseReactionsDialog, () => IsEtabs);
            GetModalMassParticipationRatiosCommand = new RelayCommand(OpenModalMassParticipationRatiosDialog, () => IsEtabs);
            GetStoryForcesCommand = new RelayCommand(OpenStoryForcesDialog, () => IsEtabs);
            GetStoryDriftsCommand = new RelayCommand(OpenStoryDriftsDialog, () => IsEtabs);
            GetStoryMaxOverAverageDisplacementsCommand = new RelayCommand(OpenStoryMaxOverAverageDisplacementsDialog, () => IsEtabs);
            GetStoryMaxOverAverageDriftsCommand = new RelayCommand(OpenStoryMaxOverAverageDriftsDialog, () => IsEtabs);
            GetMassSummaryByStoryCommand = new RelayCommand(OpenMassSummaryByStoryDialog, () => IsEtabs);
            
            GetFrameSectionsCommand = new RelayCommand(GetFrameSections, () => IsConnected);
            EditFrameSectionCommand = new RelayCommand<CSISapModelFrameSectionDTO>(EditFrameSection, _ => IsConnected);

            CurrentModelUnitText = "Not yet attached";
            SetTableGroup("ANALYSIS RESULTS", "Base Reactions");
            LoadConnectionState(showMessage: false);
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
                RefreshCommandStates();
            }
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
            get { return _currentModelUnitText; }
            private set
            {
                _currentModelUnitText = value;
                OnPropertyChanged();
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
                    return;
                }

                _selectedUnitSystem = value;
                OnPropertyChanged();
                if (value != null)
                {
                    CurrentModelUnitText = value.PresentUnitsText;
                }

                _etabsUnitService.SetSelectedUnitSystem(
                    value == null ? null : value.ToDto(),
                    value == null ? null : value.DisplayName,
                    value == null ? null : value.PresentUnitsText);

                if (!_isInitializingUnitSystems && IsConnected)
                {
                    ApplySelectedGlobalUnit(showMessages: true);
                }
            }
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

        public string ProductTitle => $"{_productName} Toolbox";

        public bool IsEtabs => string.Equals(_productName, "ETABS", StringComparison.OrdinalIgnoreCase);

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
                    default: return "Section Property";
                }
            }
        }

        public string ActivePageBreadcrumb
        {
            get
            {
                if (ActiveWorkspacePage == 7)
                {
                    return "ETABS Toolbox / General Information / Section Property - Stiffness Modifier";
                }

                if (ActiveWorkspacePage == 8)
                {
                    return "ETABS Toolbox / MODELLING HELPER / Helpers";
                }

                return ActiveWorkspacePage == 6
                    ? $"ETABS Toolbox / {ActiveTableCategory} / {ActivePageTitle}"
                    : $"ETABS Toolbox / {ActivePageTitle}";
            }
        }

        public string ActiveTableCategory
        {
            get
            {
                return string.IsNullOrWhiteSpace(_activeTableCategory)
                    ? "ANALYSIS RESULTS"
                    : _activeTableCategory;
            }
            private set
            {
                if (_activeTableCategory == value)
                {
                    return;
                }

                _activeTableCategory = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(ActivePageBreadcrumb));
            }
        }

        public string ActiveAnalysisResultsGroup
        {
            get { return _activeAnalysisResultsGroup; }
            private set
            {
                if (_activeAnalysisResultsGroup == value)
                {
                    return;
                }

                _activeAnalysisResultsGroup = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(ActivePageTitle));
                OnPropertyChanged(nameof(ActivePageBreadcrumb));
            }
        }

        public string SelectedAnalysisResultTable
        {
            get { return _selectedAnalysisResultTable; }
            set
            {
                if (_selectedAnalysisResultTable == value)
                {
                    return;
                }

                _selectedAnalysisResultTable = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(AnalysisResultPlaceholderText));
            }
        }

        public string AnalysisResultPlaceholderText
        {
            get
            {
                return string.IsNullOrWhiteSpace(SelectedAnalysisResultTable)
                    ? "Select an ETABS result table."
                    : SelectedAnalysisResultTable;
            }
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

        public System.Collections.ObjectModel.ObservableCollection<ExcelCSIToolBox.Data.DTOs.CSI.CSISapModelLoadPatternDTO> LoadPatterns { get; }
        public System.Collections.ObjectModel.ObservableCollection<ExcelCSIToolBox.Data.DTOs.CSI.CSISapModelLoadCombinationDTO> LoadCombinations { get; }
        public System.Collections.ObjectModel.ObservableCollection<CSISapModelFrameSectionDTO> FrameSections { get; }
        public System.Collections.ObjectModel.ObservableCollection<SectionDimensionAnnotation> SectionDimensionAnnotations { get; }
        public System.Collections.ObjectModel.ObservableCollection<AnalysisResultItem> AnalysisResultTables { get; }
        public System.Collections.ObjectModel.ObservableCollection<object> EtabsTableItems { get; }

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
