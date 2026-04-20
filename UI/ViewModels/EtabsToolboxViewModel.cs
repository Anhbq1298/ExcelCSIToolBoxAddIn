using System.Windows;
using System.Windows.Input;
using ExcelCSIToolBoxAddIn.Common.Commands;
using ExcelCSIToolBoxAddIn.Common.Results;
using ExcelCSIToolBoxAddIn.Core.Application;
using ExcelCSIToolBoxAddIn.Infrastructure.Etabs;
using ExcelCSIToolBoxAddIn.Infrastructure.Excel;

namespace ExcelCSIToolBoxAddIn.UI.ViewModels
{
    /// <summary>
    /// ViewModel for ETABS toolbox shell.
    /// Exposes connection state, model name, and point/frame placeholder commands.
    /// </summary>
    public class EtabsToolboxViewModel : ViewModelBase
    {
        private readonly LoadEtabsConnectionUseCase _loadEtabsConnectionUseCase;
        private readonly CloseCurrentEtabsInstanceUseCase _closeCurrentEtabsInstanceUseCase;
        private readonly GetSelectedEtabsPointsUseCase _getSelectedEtabsPointsUseCase;
        private readonly GetSelectedEtabsFramesUseCase _getSelectedEtabsFramesUseCase;
        private readonly SelectPointsFromExcelRangeByUniqueNameUseCase _selectPointsFromExcelRangeByUniqueNameUseCase;
        private readonly SelectFramesFromExcelRangeByUniqueNameUseCase _selectFramesFromExcelRangeByUniqueNameUseCase;
        private readonly AddPointsFromExcelRangeUseCase _addPointsFromExcelRangeUseCase;
        private readonly AddFrameByCoordinatesFromExcelRangeUseCase _addFrameByCoordinatesFromExcelRangeUseCase;
        private readonly AddFramesByPointFromExcelRangeUseCase _addFramesByPointFromExcelRangeUseCase;

        private readonly CreateSteelISectionsFromExcelRangeUseCase _createSteelISectionsUseCase;
        private readonly CreateSteelChannelSectionsFromExcelRangeUseCase _createSteelChannelSectionsUseCase;
        private readonly CreateSteelAngleSectionsFromExcelRangeUseCase _createSteelAngleSectionsUseCase;
        private readonly CreateSteelPipeSectionsFromExcelRangeUseCase _createSteelPipeSectionsUseCase;
        private readonly CreateSteelTubeSectionsFromExcelRangeUseCase _createSteelTubeSectionsUseCase;

        private readonly CreateConcreteRectangleSectionsFromExcelRangeUseCase _createConcreteRectangleSectionsUseCase;
        private readonly CreateConcreteCircleSectionsFromExcelRangeUseCase _createConcreteCircleSectionsUseCase;

        private string _modelName;
        private bool _isConnected;
        private string _statusText;
        private string _currentModelUnitText;
        private string _modelPath;

        public EtabsToolboxViewModel(
            IEtabsConnectionService etabsConnectionService,
            IExcelSelectionService excelSelectionService,
            IExcelOutputService excelOutputService)
        {
            _loadEtabsConnectionUseCase = new LoadEtabsConnectionUseCase(etabsConnectionService);
            _closeCurrentEtabsInstanceUseCase = new CloseCurrentEtabsInstanceUseCase(etabsConnectionService);
            _getSelectedEtabsPointsUseCase = new GetSelectedEtabsPointsUseCase(etabsConnectionService, excelOutputService);
            _getSelectedEtabsFramesUseCase = new GetSelectedEtabsFramesUseCase(etabsConnectionService, excelOutputService);
            _selectPointsFromExcelRangeByUniqueNameUseCase = new SelectPointsFromExcelRangeByUniqueNameUseCase(etabsConnectionService, excelSelectionService);
            _selectFramesFromExcelRangeByUniqueNameUseCase = new SelectFramesFromExcelRangeByUniqueNameUseCase(etabsConnectionService, excelSelectionService);
            _addPointsFromExcelRangeUseCase = new AddPointsFromExcelRangeUseCase(etabsConnectionService, excelSelectionService);
            _addFrameByCoordinatesFromExcelRangeUseCase = new AddFrameByCoordinatesFromExcelRangeUseCase(etabsConnectionService, excelSelectionService);
            _addFramesByPointFromExcelRangeUseCase = new AddFramesByPointFromExcelRangeUseCase(etabsConnectionService, excelSelectionService);
            _createSteelISectionsUseCase = new CreateSteelISectionsFromExcelRangeUseCase(etabsConnectionService, excelSelectionService);
            _createSteelChannelSectionsUseCase = new CreateSteelChannelSectionsFromExcelRangeUseCase(etabsConnectionService, excelSelectionService);
            _createSteelAngleSectionsUseCase = new CreateSteelAngleSectionsFromExcelRangeUseCase(etabsConnectionService, excelSelectionService);
            _createSteelPipeSectionsUseCase = new CreateSteelPipeSectionsFromExcelRangeUseCase(etabsConnectionService, excelSelectionService);
            _createSteelTubeSectionsUseCase = new CreateSteelTubeSectionsFromExcelRangeUseCase(etabsConnectionService, excelSelectionService);

            _createConcreteRectangleSectionsUseCase = new CreateConcreteRectangleSectionsFromExcelRangeUseCase(etabsConnectionService, excelSelectionService);
            _createConcreteCircleSectionsUseCase = new CreateConcreteCircleSectionsFromExcelRangeUseCase(etabsConnectionService, excelSelectionService);

            AttachToRunningEtabsCommand = new RelayCommand(() => LoadConnectionState(showMessage: true));
            CloseCurrentEtabsInstanceCommand = new RelayCommand(CloseCurrentEtabsInstance);

            CreateIshapeSectionCommand = new RelayCommand(() => ShowOperationResult(_createSteelISectionsUseCase.Execute()));
            CreateChannelSectionCommand = new RelayCommand(() => ShowOperationResult(_createSteelChannelSectionsUseCase.Execute()));
            CreateAngleSectionCommand = new RelayCommand(() => ShowOperationResult(_createSteelAngleSectionsUseCase.Execute()));
            CreateTubeSectionCommand = new RelayCommand(() => ShowOperationResult(_createSteelTubeSectionsUseCase.Execute()));
            CreatePipeSectionCommand = new RelayCommand(() => ShowOperationResult(_createSteelPipeSectionsUseCase.Execute()));

            CreateConcreteRectangleSectionCommand = new RelayCommand(() => ShowOperationResult(_createConcreteRectangleSectionsUseCase.Execute()));
            CreateConcreteCircleSectionCommand = new RelayCommand(() => ShowOperationResult(_createConcreteCircleSectionsUseCase.Execute()));

            SelectPointsByUniqueNameCommand = new RelayCommand(SelectPointsByUniqueName);
            SelectFramesByUniqueNameCommand = new RelayCommand(SelectFramesByUniqueName);
            AddPointByCartesianCommand = new RelayCommand(AddPointByCartesian);
            SetPointsCommand = new RelayCommand(() => ShowPlaceholder("Set Points"));
            RenameSelectedPointsCommand = new RelayCommand(() => ShowPlaceholder("Rename Selected Points"));
            GetSelectedPointsCommand = new RelayCommand(GetSelectedPoints);

            AddFramesByCoordinatesCommand = new RelayCommand(AddFramesByCoordinates);
            AddFramesByPointNamesCommand = new RelayCommand(AddFramesByPointNames);
            SetFramesCommand = new RelayCommand(() => ShowPlaceholder("Set Frames"));
            RenameFramesCommand = new RelayCommand(() => ShowPlaceholder("Rename Frames"));
            GetSelectedFramesCommand = new RelayCommand(GetSelectedFrames);
            GetFrameSectionPropertyCommand = new RelayCommand(() => ShowPlaceholder("Get Frame Section Property"));
            SetFrameSectionPropertyCommand = new RelayCommand(() => ShowPlaceholder("Set Frame Section Property"));
            GetFrameGroupAssignmentCommand = new RelayCommand(() => ShowPlaceholder("Get Frame Group Assignment"));
            SetFrameGroupAssignmentCommand = new RelayCommand(() => ShowPlaceholder("Set Frame Group Assignment"));
            GetFrameModifierCommand = new RelayCommand(() => ShowPlaceholder("Get Frame Modifier"));
            SetFrameModifierCommand = new RelayCommand(() => ShowPlaceholder("Set Frame Modifier"));
            GetPointGroupAssignmentCommand = new RelayCommand(() => ShowPlaceholder("Get Point Group Assignment"));
            SetPointGroupAssignmentCommand = new RelayCommand(() => ShowPlaceholder("Set Point Group Assignment"));

            CurrentModelUnitText = "Not yet attached";
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

        public string CurrentModelUnitText
        {
            get { return _currentModelUnitText; }
            private set
            {
                _currentModelUnitText = value;
                OnPropertyChanged();
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

        public ICommand AttachToRunningEtabsCommand { get; }
        public ICommand CloseCurrentEtabsInstanceCommand { get; }

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
        public ICommand GetPointGroupAssignmentCommand { get; }
        public ICommand SetPointGroupAssignmentCommand { get; }

        private void LoadConnectionState(bool showMessage)
        {
            var result = _loadEtabsConnectionUseCase.Execute();

            if (result.IsSuccess && result.Data != null)
            {
                IsConnected = true;
                ModelName = string.IsNullOrWhiteSpace(result.Data.ModelFileName)
                    ? "Unknown model"
                    : result.Data.ModelFileName;
                ModelPath = result.Data.ModelPath ?? string.Empty;
                CurrentModelUnitText = string.IsNullOrWhiteSpace(result.Data.ModelCurrentUnit)
                    ? "Units unavailable"
                    : result.Data.ModelCurrentUnit;

                StatusText = "Connected to running ETABS instance.";

                if (showMessage)
                {
                    MessageBox.Show(
                        "Successfully attached to running ETABS.",
                        "ETABS Toolbox",
                        MessageBoxButton.OK,
                        MessageBoxImage.Information);
                }

                return;
            }

            IsConnected = false;
            SetDetachedModelInfo("Not yet attached");
            StatusText = string.IsNullOrWhiteSpace(result.Message)
                ? "ETABS connection unavailable."
                : result.Message;

            if (showMessage)
            {
                MessageBox.Show(
                    StatusText,
                    "ETABS Toolbox",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
            }
        }

        private void CloseCurrentEtabsInstance()
        {
            var result = _closeCurrentEtabsInstanceUseCase.Execute();

            if (result.IsSuccess)
            {
                IsConnected = false;
                SetDetachedModelInfo("Not connected");
                StatusText = result.Message;

                MessageBox.Show(
                    result.Message,
                    "ETABS Toolbox",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
                return;
            }

            MessageBox.Show(
                result.Message,
                "ETABS Toolbox",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);

            StatusText = result.Message;
        }

        private void SetDetachedModelInfo(string modelNameText)
        {
            ModelName = modelNameText;
            ModelPath = string.Empty;
            CurrentModelUnitText = "Not yet attached";
        }

        private void SelectPointsByUniqueName()
        {
            ShowOperationResult(_selectPointsFromExcelRangeByUniqueNameUseCase.Execute());
        }

        private void SelectFramesByUniqueName()
        {
            ShowOperationResult(_selectFramesFromExcelRangeByUniqueNameUseCase.Execute());
        }

        private void AddPointByCartesian()
        {
            ShowOperationResult(_addPointsFromExcelRangeUseCase.Execute());
        }

        private void AddFramesByCoordinates()
        {
            ShowOperationResult(_addFrameByCoordinatesFromExcelRangeUseCase.Execute());
        }

        private void AddFramesByPointNames()
        {
            ShowOperationResult(_addFramesByPointFromExcelRangeUseCase.Execute());
        }

        private void GetSelectedPoints()
        {
            ShowOperationResult(_getSelectedEtabsPointsUseCase.Execute());
        }

        private void GetSelectedFrames()
        {
            ShowOperationResult(_getSelectedEtabsFramesUseCase.Execute());
        }

        private static void ShowOperationResult(OperationResult result)
        {
            MessageBox.Show(
                result.Message,
                "ETABS Toolbox",
                MessageBoxButton.OK,
                result.IsSuccess ? MessageBoxImage.Information : MessageBoxImage.Warning);
        }

        private static void ShowPlaceholder(string featureName)
        {
            MessageBox.Show(
                $"{featureName} is a placeholder for phase 1.",
                "ETABS Toolbox",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
        }
    }
}
