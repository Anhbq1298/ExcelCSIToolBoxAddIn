using System.Windows;
using System.Windows.Input;
using ExcelCSIToolBoxAddIn.Common.Commands;
using ExcelCSIToolBoxAddIn.Core.Application;
using ExcelCSIToolBoxAddIn.Infrastructure.Etabs;

namespace ExcelCSIToolBoxAddIn.UI.ViewModels
{
    /// <summary>
    /// ViewModel for ETABS toolbox shell.
    /// Exposes connection state, model name, and point/frame placeholder commands.
    /// </summary>
    public class EtabsToolboxViewModel : ViewModelBase
    {
        private readonly LoadEtabsConnectionUseCase _loadEtabsConnectionUseCase;

        private string _modelName;
        private bool _isConnected;
        private string _statusText;

        public EtabsToolboxViewModel(IEtabsConnectionService etabsConnectionService)
        {
            _loadEtabsConnectionUseCase = new LoadEtabsConnectionUseCase(etabsConnectionService);

            AttachToRunningEtabsCommand = new RelayCommand(LoadConnectionState);

            AddPointsCommand = new RelayCommand(() => ShowPlaceholder("Add Points"));
            SetPointsCommand = new RelayCommand(() => ShowPlaceholder("Set Points"));
            RenamePointsCommand = new RelayCommand(() => ShowPlaceholder("Rename Points"));
            GetPointsCommand = new RelayCommand(() => ShowPlaceholder("Get Points"));

            AddFramesCommand = new RelayCommand(() => ShowPlaceholder("Add Frames"));
            SetFramesCommand = new RelayCommand(() => ShowPlaceholder("Set Frames"));
            RenameFramesCommand = new RelayCommand(() => ShowPlaceholder("Rename Frames"));
            GetFramesCommand = new RelayCommand(() => ShowPlaceholder("Get Frames"));

            LoadConnectionState();
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

        public string ModelDisplayText => $"ETABS Model: {ModelName}";

        public ICommand AttachToRunningEtabsCommand { get; }

        public ICommand AddPointsCommand { get; }
        public ICommand SetPointsCommand { get; }
        public ICommand RenamePointsCommand { get; }
        public ICommand GetPointsCommand { get; }

        public ICommand AddFramesCommand { get; }
        public ICommand SetFramesCommand { get; }
        public ICommand RenameFramesCommand { get; }
        public ICommand GetFramesCommand { get; }

        private void LoadConnectionState()
        {
            var result = _loadEtabsConnectionUseCase.Execute();

            if (result.IsSuccess && result.Data != null)
            {
                IsConnected = true;
                ModelName = string.IsNullOrWhiteSpace(result.Data.ModelFileName)
                    ? "Unknown model"
                    : result.Data.ModelFileName;

                StatusText = "Connected to running ETABS instance.";
                return;
            }

            IsConnected = false;
            ModelName = "Not connected";
            StatusText = string.IsNullOrWhiteSpace(result.Message)
                ? "ETABS connection unavailable."
                : result.Message;
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
