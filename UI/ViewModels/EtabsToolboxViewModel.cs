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
        private readonly GetCurrentEtabsModelUnitsUseCase _getCurrentEtabsModelUnitsUseCase;
        private readonly CloseCurrentEtabsInstanceUseCase _closeCurrentEtabsInstanceUseCase;
        private readonly GetSelectedEtabsPointsUseCase _getSelectedEtabsPointsUseCase;
        private readonly SelectPointsFromExcelRangeByUniqueNameUseCase _selectPointsFromExcelRangeByUniqueNameUseCase;
        private readonly AddPointsFromExcelRangeUseCase _addPointsFromExcelRangeUseCase;

        private string _modelName;
        private bool _isConnected;
        private string _statusText;
        private string _currentModelUnitText;

        public EtabsToolboxViewModel(
            IEtabsConnectionService etabsConnectionService,
            IExcelOutputService excelOutputService)
        {
            var excelSelectionService = new ExcelSelectionService();

            _loadEtabsConnectionUseCase = new LoadEtabsConnectionUseCase(etabsConnectionService);
            _getCurrentEtabsModelUnitsUseCase = new GetCurrentEtabsModelUnitsUseCase(etabsConnectionService);
            _closeCurrentEtabsInstanceUseCase = new CloseCurrentEtabsInstanceUseCase(etabsConnectionService);
            _getSelectedEtabsPointsUseCase = new GetSelectedEtabsPointsUseCase(etabsConnectionService, excelOutputService);
            _selectPointsFromExcelRangeByUniqueNameUseCase = new SelectPointsFromExcelRangeByUniqueNameUseCase(etabsConnectionService, excelSelectionService);
            _addPointsFromExcelRangeUseCase = new AddPointsFromExcelRangeUseCase(etabsConnectionService, excelSelectionService);

            AttachToRunningEtabsCommand = new RelayCommand(() => LoadConnectionState(showMessage: true));
            CloseCurrentEtabsInstanceCommand = new RelayCommand(CloseCurrentEtabsInstance);

            SelectPointsByUniqueNameCommand = new RelayCommand(SelectPointsByUniqueName);
            AddPointByCartesianCommand = new RelayCommand(AddPointByCartesian);
            SetPointsCommand = new RelayCommand(() => ShowPlaceholder("Set Points"));
            RenameSelectedPointsCommand = new RelayCommand(() => ShowPlaceholder("Rename Selected Points"));
            GetSelectedPointsCommand = new RelayCommand(GetSelectedPoints);

            AddFramesCommand = new RelayCommand(() => ShowPlaceholder("Add Frames"));
            SetFramesCommand = new RelayCommand(() => ShowPlaceholder("Set Frames"));
            RenameFramesCommand = new RelayCommand(() => ShowPlaceholder("Rename Frames"));
            GetFramesCommand = new RelayCommand(() => ShowPlaceholder("Get Frames"));

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

        public string ModelDisplayText => $"{ModelName}";

        public ICommand AttachToRunningEtabsCommand { get; }
        public ICommand CloseCurrentEtabsInstanceCommand { get; }

        public ICommand SelectPointsByUniqueNameCommand { get; }
        public ICommand AddPointByCartesianCommand { get; }
        public ICommand SetPointsCommand { get; }
        public ICommand RenameSelectedPointsCommand { get; }
        public ICommand GetSelectedPointsCommand { get; }

        public ICommand AddFramesCommand { get; }
        public ICommand SetFramesCommand { get; }
        public ICommand RenameFramesCommand { get; }
        public ICommand GetFramesCommand { get; }

        private void LoadConnectionState(bool showMessage)
        {
            var result = _loadEtabsConnectionUseCase.Execute();

            if (result.IsSuccess && result.Data != null)
            {
                IsConnected = true;
                ModelName = string.IsNullOrWhiteSpace(result.Data.ModelFileName)
                    ? "Unknown model"
                    : result.Data.ModelFileName;
                CurrentModelUnitText = GetCurrentModelUnitsTextOrFallback();

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
            ModelName = "Not yet attached";
            CurrentModelUnitText = "Not yet attached";
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
                ModelName = "Not connected";
                CurrentModelUnitText = "Not yet attached";
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

        private string GetCurrentModelUnitsTextOrFallback()
        {
            var unitResult = _getCurrentEtabsModelUnitsUseCase.Execute();
            if (unitResult.IsSuccess && !string.IsNullOrWhiteSpace(unitResult.Data))
            {
                return unitResult.Data;
            }

            return "Units unavailable";
        }

        private void SelectPointsByUniqueName()
        {
            ShowOperationResult(_selectPointsFromExcelRangeByUniqueNameUseCase.Execute());
        }

        private void AddPointByCartesian()
        {
            ShowOperationResult(_addPointsFromExcelRangeUseCase.Execute());
        }

        private void GetSelectedPoints()
        {
            var result = _getSelectedEtabsPointsUseCase.Execute();

            if (result.IsSuccess)
            {
                MessageBox.Show(
                    "Successfully exported selected ETABS points to Excel.",
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
