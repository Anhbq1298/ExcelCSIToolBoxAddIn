using System.Collections.ObjectModel;
using System.Windows.Input;
using CsiExcelAddin.Commands;
using CsiExcelAddin.Models;
using CsiExcelAddin.Services.Interfaces;

namespace CsiExcelAddin.ViewModels
{
    /// <summary>
    /// ViewModel for the main popup window.
    /// Orchestrates attach/detach, data reading from CSI, and export to Excel.
    /// All operations route through injected service interfaces â€” never direct API calls.
    /// </summary>
    public class MainViewModel : BaseViewModel
    {
        private readonly ICsiProductAdapter _adapter;
        private readonly IExcelRangeWriter _excelWriter;

        // â”€â”€ Bindable state â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

        private string _statusMessage = "Ready.";
        /// <summary>Shown in the status bar at the bottom of the popup.</summary>
        public string StatusMessage
        {
            get => _statusMessage;
            set => SetProperty(ref _statusMessage, value);
        }

        private bool _isAttached;
        /// <summary>Controls Attach/Detach button states and dependent command guards.</summary>
        public bool IsAttached
        {
            get => _isAttached;
            set => SetProperty(ref _isAttached, value);
        }

        private string _modelFileName = "â€”";
        public string ModelFileName
        {
            get => _modelFileName;
            set => SetProperty(ref _modelFileName, value);
        }

        private string _currentUnits = "â€”";
        public string CurrentUnits
        {
            get => _currentUnits;
            set => SetProperty(ref _currentUnits, value);
        }

        /// <summary>Frame sections loaded from the CSI model, bound to a DataGrid.</summary>
        public ObservableCollection<FrameSectionDto> FrameSections { get; } = new ObservableCollection<FrameSectionDto>();

        // â”€â”€ Commands â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

        public ICommand AttachCommand { get; }
        public ICommand DetachCommand { get; }
        public ICommand ReadModelCommand { get; }
        public ICommand ExportToExcelCommand { get; }

        // â”€â”€ Constructor â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

        /// <summary>
        /// Services are injected through the constructor (composition root wires them up).
        /// The ViewModel never creates service instances directly.
        /// </summary>
        public MainViewModel(ICsiProductAdapter adapter, IExcelRangeWriter excelWriter)
        {
            _adapter = adapter;
            _excelWriter = excelWriter;

            AttachCommand = new RelayCommand(ExecuteAttach, _ => !IsAttached);
            DetachCommand = new RelayCommand(ExecuteDetach, _ => IsAttached);
            ReadModelCommand = new AsyncRelayCommand(ExecuteReadModelAsync, () => IsAttached);
            ExportToExcelCommand = new RelayCommand(ExecuteExportToExcel, _ => FrameSections.Count > 0);
        }

        // â”€â”€ Command handlers â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

        private void ExecuteAttach(object _)
        {
            var result = _adapter.ApplicationService.Attach();
            StatusMessage = result.Message;

            if (result.Success)
            {
                IsAttached = true;
                ModelFileName = _adapter.ModelService.GetModelFileName();
                CurrentUnits = _adapter.ModelService.GetCurrentUnits();
            }
        }

        private void ExecuteDetach(object _)
        {
            _adapter.ApplicationService.Detach();
            IsAttached = false;
            ModelFileName = "â€”";
            CurrentUnits = "â€”";
            FrameSections.Clear();
            StatusMessage = "Detached from " + _adapter.ProductName + ".";
        }

        private async System.Threading.Tasks.Task ExecuteReadModelAsync()
        {
            StatusMessage = "Reading model data...";
            FrameSections.Clear();

            // Offload to a background thread so Excel UI remains responsive
            var sections = await System.Threading.Tasks.Task.Run(
                () => _adapter.ModelService.GetFrameSections());

            foreach (var section in sections)
                FrameSections.Add(section);

            StatusMessage = $"Loaded {FrameSections.Count} frame sections.";
        }

        private void ExecuteExportToExcel(object _)
        {
            // Build rows: header + data
            var rows = new System.Collections.Generic.List<System.Collections.Generic.IReadOnlyList<object>>();
            rows.Add(new[] { "Name", "Material", "Depth", "Width" });

            foreach (var s in FrameSections)
                rows.Add(new object[] { s.Name, s.Material, s.Depth, s.Width });

            // Write to a named range defined in the workbook ("CsiFrameSections")
            _excelWriter.WriteToNamedRange("CsiFrameSections", rows);
            StatusMessage = $"Exported {FrameSections.Count} sections to Excel.";
        }
    }
}

