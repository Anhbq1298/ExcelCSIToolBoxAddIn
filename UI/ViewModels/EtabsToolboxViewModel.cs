using System.Windows;
using System.Windows.Input;
using ExcelCSIToolBoxAddIn.Common.Commands;
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
        private readonly GetSelectedEtabsPointsUseCase _getSelectedEtabsPointsUseCase;

        private string _modelName;
        private bool _isConnected;
        private string _statusText;

        public EtabsToolboxViewModel(
            IEtabsConnectionService etabsConnectionService,
            IExcelOutputService excelOutputService)
        {
            _loadEtabsConnectionUseCase = new LoadEtabsConnectionUseCase(etabsConnectionService);
            _getSelectedEtabsPointsUseCase = new GetSelectedEtabsPointsUseCase(etabsConnectionService, excelOutputService);

            AttachToRunningEtabsCommand = new RelayCommand(() => LoadConnectionState(showMessage: true));

            AddPointsCommand = new RelayCommand(() => ShowPlaceholder("Add Points"));
            SetPointsCommand = new RelayCommand(() => ShowPlaceholder("Set Points"));
            RenameSelectedPointsCommand = new RelayCommand(() => ShowPlaceholder("Rename Selected Points"));
            GetSelectedPointsCommand = new RelayCommand(GetSelectedPoints);

            AddFramesCommand = new RelayCommand(() => ShowPlaceholder("Add Frames"));
            SetFramesCommand = new RelayCommand(() => ShowPlaceholder("Set Frames"));
            RenameFramesCommand = new RelayCommand(() => ShowPlaceholder("Rename Frames"));
            GetFramesCommand = new RelayCommand(() => ShowPlaceholder("Get Frames"));

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

        public string ModelDisplayText => $"{ModelName}";

        public ICommand AttachToRunningEtabsCommand { get; }

        public ICommand AddPointsCommand { get; }
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
            ModelName = "Not connected";
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
