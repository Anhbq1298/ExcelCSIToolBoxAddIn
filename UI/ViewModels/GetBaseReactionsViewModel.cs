using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Input;
using ExcelCSIToolBox.Application.UseCases;
using ExcelCSIToolBox.Core.Abstractions.CSI;
using ExcelCSIToolBox.Core.Abstractions.Excel;
using ExcelCSIToolBox.Core.Common.Commands;
using ExcelCSIToolBox.Core.Common.Results;
using ExcelCSIToolBox.Data.DTOs.CSI;
using ExcelCSIToolBox.Infrastructure.Excel;
using ExcelRange = Microsoft.Office.Interop.Excel.Range;

namespace ExcelCSIToolBoxAddIn.UI.ViewModels
{
    public class GetBaseReactionsViewModel : ViewModelBase
    {
        private readonly ICSISapModelConnectionService _csiConnectionService;
        private readonly IExcelOutputService _excelOutputService;
        private readonly GetBaseReactionsUseCase _getBaseReactionsUseCase;
        private string _anchorCellAddress;
        private string _statusText;
        private bool _isBusy;

        public GetBaseReactionsViewModel(
            CsiToolboxUseCaseBundle useCases,
            ICSISapModelConnectionService csiConnectionService,
            IExcelOutputService excelOutputService)
        {
            if (useCases == null) throw new ArgumentNullException(nameof(useCases));
            _csiConnectionService = csiConnectionService ?? throw new ArgumentNullException(nameof(csiConnectionService));
            _excelOutputService = excelOutputService ?? throw new ArgumentNullException(nameof(excelOutputService));
            _getBaseReactionsUseCase = useCases.GetBaseReactions ?? throw new ArgumentNullException(nameof(useCases.GetBaseReactions));

            OutputCases = new ObservableCollection<BaseReactionOutputCaseViewModel>();
            LoadOutputCasesCommand = new RelayCommand(LoadOutputCases, () => !IsBusy);
            UseActiveCellCommand = new RelayCommand(UseActiveCell, () => !IsBusy);
            PickAnchorCellCommand = new RelayCommand(PickAnchorCell, () => !IsBusy);
            RunCommand = new RelayCommand(Run, () => !IsBusy);
            CancelCommand = new RelayCommand(() => RequestClose?.Invoke(this, EventArgs.Empty));

            UseActiveCell();
            LoadOutputCases();
        }

        public event EventHandler RequestClose;

        public ObservableCollection<BaseReactionOutputCaseViewModel> OutputCases { get; private set; }

        public ICommand LoadOutputCasesCommand { get; private set; }
        public ICommand UseActiveCellCommand { get; private set; }
        public ICommand PickAnchorCellCommand { get; private set; }
        public ICommand RunCommand { get; private set; }
        public ICommand CancelCommand { get; private set; }

        public string AnchorCellAddress
        {
            get { return _anchorCellAddress; }
            private set
            {
                _anchorCellAddress = value;
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

        public bool IsBusy
        {
            get { return _isBusy; }
            private set
            {
                _isBusy = value;
                OnPropertyChanged();
                RaiseCommandStates();
            }
        }

        private void LoadOutputCases()
        {
            if (!EnsureEtabs())
            {
                return;
            }

            try
            {
                IsBusy = true;
                StatusText = "Loading ETABS load cases and combinations...";
                var result = _csiConnectionService.GetAnalysisOutputCases();
                if (!result.IsSuccess)
                {
                    ShowWarning(result.Message);
                    StatusText = result.Message;
                    return;
                }

                OutputCases.Clear();
                if (result.Data != null)
                {
                    foreach (var outputCase in result.Data)
                    {
                        OutputCases.Add(new BaseReactionOutputCaseViewModel(outputCase));
                    }
                }

                StatusText = OutputCases.Count == 0
                    ? "No ETABS load cases or load combinations were found."
                    : $"Loaded {OutputCases.Count} ETABS load case(s) / combination(s).";
            }
            catch (Exception ex)
            {
                StatusText = "Failed to load ETABS output cases.";
                ShowError($"Failed to load ETABS load cases and combinations: {ex.Message}");
            }
            finally
            {
                IsBusy = false;
            }
        }

        private void UseActiveCell()
        {
            ExcelRange activeCell = GetActiveExcelCell();
            if (activeCell == null)
            {
                AnchorCellAddress = string.Empty;
                StatusText = "Select an Excel anchor cell for the first data row.";
                return;
            }

            AnchorCellAddress = FormatAddress(activeCell);
            StatusText = $"Anchor cell set to {AnchorCellAddress}.";
        }

        private void PickAnchorCell()
        {
            try
            {
                var excelApp = ExcelApplicationProvider.GetApplication();
                if (excelApp == null)
                {
                    ShowWarning("Excel application is not available.");
                    return;
                }

                object result = excelApp.InputBox(
                    "Select the top-left anchor cell where the first Base Reactions data row should start. Headers are excluded.",
                    "Get Base Reactions",
                    Type: 8);

                if (result is bool && (bool)result == false)
                {
                    return;
                }

                var selectedRange = result as ExcelRange;
                ExcelRange startCell = selectedRange == null ? null : selectedRange.Cells[1, 1] as ExcelRange;
                if (startCell == null)
                {
                    ShowWarning("No Excel anchor cell was selected.");
                    return;
                }

                startCell.Select();
                AnchorCellAddress = FormatAddress(startCell);
                StatusText = $"Anchor cell set to {AnchorCellAddress}.";
            }
            catch (Exception ex)
            {
                ShowError($"Failed to select the Excel anchor cell: {ex.Message}");
            }
        }

        private void Run()
        {
            if (!EnsureEtabs())
            {
                return;
            }

            if (string.IsNullOrWhiteSpace(AnchorCellAddress))
            {
                ShowWarning("Select or confirm the Excel anchor cell before running.");
                return;
            }

            var selectedCases = GetSelectedOutputCases();
            if (selectedCases.Count == 0)
            {
                ShowWarning("Select at least one ETABS load case or load combination.");
                return;
            }

            try
            {
                IsBusy = true;
                StatusText = "Extracting ETABS Base Reactions...";
                var result = _getBaseReactionsUseCase.Execute(selectedCases);
                if (!result.IsSuccess)
                {
                    StatusText = result.Message;
                    ShowWarning(result.Message);
                    return;
                }

                if (result.Data == null || result.Data.Count == 0)
                {
                    StatusText = "ETABS returned no Base Reactions records.";
                    MessageBox.Show(
                        "ETABS returned no Base Reactions records for the selected cases/combinations. Nothing was written to Excel.",
                        "Get Base Reactions",
                        MessageBoxButton.OK,
                        MessageBoxImage.Information);
                    return;
                }

                object[,] values = CreateOutputValues(result.Data);
                OperationResult writeResult = _excelOutputService.WriteValuesToActiveCell(
                    values,
                    $"Successfully wrote {result.Data.Count} Base Reactions record(s) to Excel.");

                StatusText = writeResult.Message;
                MessageBox.Show(
                    writeResult.Message,
                    "Get Base Reactions",
                    MessageBoxButton.OK,
                    writeResult.IsSuccess ? MessageBoxImage.Information : MessageBoxImage.Warning);
            }
            catch (Exception ex)
            {
                StatusText = "Failed to extract Base Reactions.";
                ShowError($"Failed to extract Base Reactions: {ex.Message}");
            }
            finally
            {
                IsBusy = false;
            }
        }

        private List<CSISapModelOutputCaseDTO> GetSelectedOutputCases()
        {
            var selectedCases = new List<CSISapModelOutputCaseDTO>();
            foreach (var item in OutputCases)
            {
                if (item != null && item.IsSelected && item.OutputCase != null)
                {
                    selectedCases.Add(item.OutputCase);
                }
            }

            return selectedCases;
        }

        private static object[,] CreateOutputValues(IReadOnlyList<CSISapModelBaseReactionRowDTO> rows)
        {
            var values = new object[rows.Count, 13];
            for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++)
            {
                CSISapModelBaseReactionRowDTO row = rows[rowIndex];
                values[rowIndex, 0] = row.OutputCase;
                values[rowIndex, 1] = row.CaseType;
                values[rowIndex, 2] = row.StepType;
                values[rowIndex, 3] = row.StepNumber;
                values[rowIndex, 4] = row.FX;
                values[rowIndex, 5] = row.FY;
                values[rowIndex, 6] = row.FZ;
                values[rowIndex, 7] = row.MX;
                values[rowIndex, 8] = row.MY;
                values[rowIndex, 9] = row.MZ;
                values[rowIndex, 10] = row.X;
                values[rowIndex, 11] = row.Y;
                values[rowIndex, 12] = row.Z;
            }

            return values;
        }

        private bool EnsureEtabs()
        {
            if (!string.Equals(_csiConnectionService.ProductName, "ETABS", StringComparison.OrdinalIgnoreCase))
            {
                ShowWarning("Get Base Reactions is available from the ETABS Toolbox only.");
                return false;
            }

            var connectionResult = _csiConnectionService.GetCurrentConnection();
            if (!connectionResult.IsSuccess)
            {
                ShowWarning(string.IsNullOrWhiteSpace(connectionResult.Message)
                    ? "ETABS is not attached. Attach to a running ETABS instance first."
                    : connectionResult.Message);
                return false;
            }

            return true;
        }

        private static ExcelRange GetActiveExcelCell()
        {
            try
            {
                var excelApp = ExcelApplicationProvider.GetApplication();
                if (excelApp == null)
                {
                    return null;
                }

                var selectedRange = excelApp.Selection as ExcelRange;
                if (selectedRange != null)
                {
                    return selectedRange.Cells[1, 1] as ExcelRange;
                }

                return excelApp.ActiveCell as ExcelRange;
            }
            catch
            {
                return null;
            }
        }

        private static string FormatAddress(ExcelRange cell)
        {
            if (cell == null)
            {
                return string.Empty;
            }

            string address = cell.Address[RowAbsolute: false, ColumnAbsolute: false];
            string sheetName = cell.Worksheet == null ? string.Empty : cell.Worksheet.Name;
            return string.IsNullOrWhiteSpace(sheetName) ? address : $"{sheetName}!{address}";
        }

        private void RaiseCommandStates()
        {
            RaiseCommandState(LoadOutputCasesCommand);
            RaiseCommandState(UseActiveCellCommand);
            RaiseCommandState(PickAnchorCellCommand);
            RaiseCommandState(RunCommand);
        }

        private static void RaiseCommandState(ICommand command)
        {
            var relayCommand = command as IRelayCommand;
            if (relayCommand != null)
            {
                relayCommand.RaiseCanExecuteChanged();
            }
        }

        private static void ShowWarning(string message)
        {
            MessageBox.Show(
                string.IsNullOrWhiteSpace(message) ? "The operation could not be completed." : message,
                "Get Base Reactions",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
        }

        private static void ShowError(string message)
        {
            MessageBox.Show(
                string.IsNullOrWhiteSpace(message) ? "An unexpected error occurred." : message,
                "Get Base Reactions",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
        }
    }
}
