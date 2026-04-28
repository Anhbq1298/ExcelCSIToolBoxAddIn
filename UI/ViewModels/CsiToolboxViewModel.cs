using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using ExcelCSIToolBoxAddIn.Common.Commands;
using ExcelCSIToolBoxAddIn.Common.Results;
using ExcelCSIToolBoxAddIn.Core.Application;
using ExcelCSIToolBoxAddIn.Infrastructure.CSISapModel;
using ExcelCSIToolBoxAddIn.Infrastructure.Excel;

namespace ExcelCSIToolBoxAddIn.UI.ViewModels
{
    /// <summary>
    /// ViewModel for CSI toolbox shells.
    /// Exposes connection state, model name, and point/frame placeholder commands.
    /// </summary>
    public class CsiToolboxViewModel : ViewModelBase
    {
        private readonly LoadCSISapModelConnectionUseCase _loadCSISapModelConnectionUseCase;
        private readonly CloseCurrentInstanceUseCase _closeCurrentInstanceUseCase;
        private readonly GetSelectedCSISapModelPointsUseCase _getSelectedCSISapModelPointsUseCase;
        private readonly GetSelectedCSISapModelFramesUseCase _getSelectedCSISapModelFramesUseCase;
        private readonly SelectPointsFromExcelRangeByUniqueNameUseCase _selectPointsFromExcelRangeByUniqueNameUseCase;
        private readonly SelectFramesFromExcelRangeByUniqueNameUseCase _selectFramesFromExcelRangeByUniqueNameUseCase;
        private readonly AddPointsFromExcelRangeUseCase _addPointsFromExcelRangeUseCase;
        private readonly AddFrameByCoordinatesFromExcelRangeUseCase _addFrameByCoordinatesFromExcelRangeUseCase;
        private readonly AddFramesByPointFromExcelRangeUseCase _addFramesByPointFromExcelRangeUseCase;
        private readonly CreateShellAreasFromSelectedFramesUseCase _createShellAreasFromSelectedFramesUseCase;

        private readonly CreateSteelISectionsFromExcelRangeUseCase _createSteelISectionsUseCase;
        private readonly CreateSteelChannelSectionsFromExcelRangeUseCase _createSteelChannelSectionsUseCase;
        private readonly CreateSteelAngleSectionsFromExcelRangeUseCase _createSteelAngleSectionsUseCase;
        private readonly CreateSteelPipeSectionsFromExcelRangeUseCase _createSteelPipeSectionsUseCase;
        private readonly CreateSteelTubeSectionsFromExcelRangeUseCase _createSteelTubeSectionsUseCase;

        private readonly CreateConcreteRectangleSectionsFromExcelRangeUseCase _createConcreteRectangleSectionsUseCase;
        private readonly CreateConcreteCircleSectionsFromExcelRangeUseCase _createConcreteCircleSectionsUseCase;
        
        private readonly GetLoadCombinationsUseCase _getLoadCombinationsUseCase;
        private readonly DeleteLoadCombinationsUseCase _deleteLoadCombinationsUseCase;
        private readonly GetLoadCombinationDetailsUseCase _getLoadCombinationDetailsUseCase;

        private readonly GetLoadPatternsUseCase _getLoadPatternsUseCase;
        private readonly DeleteLoadPatternsUseCase _deleteLoadPatternsUseCase;

        private string _modelName;
        private bool _isConnected;
        private string _statusText;
        private string _currentModelUnitText;
        private string _modelPath;
        private readonly string _productName;

        public CsiToolboxViewModel(
            ICSISapModelConnectionService csiConnectionService,
            IExcelSelectionService excelSelectionService,
            IExcelOutputService excelOutputService)
        {
            _productName = string.IsNullOrWhiteSpace(csiConnectionService.ProductName)
                ? "CSI"
                : csiConnectionService.ProductName;

            _loadCSISapModelConnectionUseCase = new LoadCSISapModelConnectionUseCase(csiConnectionService);
            _closeCurrentInstanceUseCase = new CloseCurrentInstanceUseCase(csiConnectionService);
            _getSelectedCSISapModelPointsUseCase = new GetSelectedCSISapModelPointsUseCase(csiConnectionService, excelOutputService);
            _getSelectedCSISapModelFramesUseCase = new GetSelectedCSISapModelFramesUseCase(csiConnectionService, excelOutputService);
            _selectPointsFromExcelRangeByUniqueNameUseCase = new SelectPointsFromExcelRangeByUniqueNameUseCase(csiConnectionService, excelSelectionService);
            _selectFramesFromExcelRangeByUniqueNameUseCase = new SelectFramesFromExcelRangeByUniqueNameUseCase(csiConnectionService, excelSelectionService);
            _addPointsFromExcelRangeUseCase = new AddPointsFromExcelRangeUseCase(csiConnectionService, excelSelectionService);
            _addFrameByCoordinatesFromExcelRangeUseCase = new AddFrameByCoordinatesFromExcelRangeUseCase(csiConnectionService, excelSelectionService);
            _addFramesByPointFromExcelRangeUseCase = new AddFramesByPointFromExcelRangeUseCase(csiConnectionService, excelSelectionService);
            _createShellAreasFromSelectedFramesUseCase = new CreateShellAreasFromSelectedFramesUseCase(csiConnectionService);
            _createSteelISectionsUseCase = new CreateSteelISectionsFromExcelRangeUseCase(csiConnectionService, excelSelectionService);
            _createSteelChannelSectionsUseCase = new CreateSteelChannelSectionsFromExcelRangeUseCase(csiConnectionService, excelSelectionService);
            _createSteelAngleSectionsUseCase = new CreateSteelAngleSectionsFromExcelRangeUseCase(csiConnectionService, excelSelectionService);
            _createSteelPipeSectionsUseCase = new CreateSteelPipeSectionsFromExcelRangeUseCase(csiConnectionService, excelSelectionService);
            _createSteelTubeSectionsUseCase = new CreateSteelTubeSectionsFromExcelRangeUseCase(csiConnectionService, excelSelectionService);

            _createConcreteRectangleSectionsUseCase = new CreateConcreteRectangleSectionsFromExcelRangeUseCase(csiConnectionService, excelSelectionService);
            _createConcreteCircleSectionsUseCase = new CreateConcreteCircleSectionsFromExcelRangeUseCase(csiConnectionService, excelSelectionService);

            _getLoadCombinationsUseCase = new GetLoadCombinationsUseCase(csiConnectionService);
            _deleteLoadCombinationsUseCase = new DeleteLoadCombinationsUseCase(csiConnectionService);
            _getLoadCombinationDetailsUseCase = new GetLoadCombinationDetailsUseCase(csiConnectionService);

            _getLoadPatternsUseCase = new GetLoadPatternsUseCase(csiConnectionService);
            _deleteLoadPatternsUseCase = new DeleteLoadPatternsUseCase(csiConnectionService);

            LoadCombinations = new System.Collections.ObjectModel.ObservableCollection<ExcelCSIToolBoxAddIn.Data.DTOs.CSISapModelLoadCombinationDTO>();
            LoadPatterns = new System.Collections.ObjectModel.ObservableCollection<ExcelCSIToolBoxAddIn.Data.DTOs.CSISapModelLoadPatternDTO>();

            AttachToRunningCsiCommand = new RelayCommand(() => LoadConnectionState(showMessage: true));
            CloseCurrentInstanceCommand = new RelayCommand(CloseCurrentInstance);

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
            CreateShellAreasFromSelectedFramesCommand = new RelayCommand(CreateShellAreasFromSelectedFrames);
            GetPointGroupAssignmentCommand = new RelayCommand(() => ShowPlaceholder("Get Point Group Assignment"));
            SetPointGroupAssignmentCommand = new RelayCommand(() => ShowPlaceholder("Set Point Group Assignment"));

            GetLoadPatternsCommand = new RelayCommand(GetLoadPatterns);
            AddLoadPatternFromExcelCommand = new RelayCommand(() => ShowPlaceholder("Add Load Pattern From Excel"));
            DeleteSelectedLoadPatternsCommand = new RelayCommand<System.Collections.IList>(DeleteSelectedLoadPatterns);
            
            GetLoadCombinationsCommand = new RelayCommand(GetLoadCombinations);
            AddLoadCombinationFromExcelCommand = new RelayCommand(() => ShowPlaceholder("Add Load Combination From Excel"));
            DeleteSelectedLoadCombinationsCommand = new RelayCommand<System.Collections.IList>(DeleteSelectedLoadCombinations);
            ViewLoadCombinationCommand = new RelayCommand<System.Collections.IList>(ViewLoadCombination);

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

        public string ProductTitle => $"{_productName} Toolbox";

        public ICommand AttachToRunningCsiCommand { get; }
        public ICommand CloseCurrentInstanceCommand { get; }

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
        public ICommand AddLoadCombinationFromExcelCommand { get; }
        public ICommand DeleteSelectedLoadCombinationsCommand { get; }
        public ICommand ViewLoadCombinationCommand { get; }

        public System.Collections.ObjectModel.ObservableCollection<ExcelCSIToolBoxAddIn.Data.DTOs.CSISapModelLoadPatternDTO> LoadPatterns { get; }
        public System.Collections.ObjectModel.ObservableCollection<ExcelCSIToolBoxAddIn.Data.DTOs.CSISapModelLoadCombinationDTO> LoadCombinations { get; }

        private void LoadConnectionState(bool showMessage)
        {
            var result = _loadCSISapModelConnectionUseCase.Execute();

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

                StatusText = $"Connected to running {_productName} instance.";

                if (showMessage)
                {
                    MessageBox.Show(
                        $"Successfully attached to running {_productName}.",
                        ProductTitle,
                        MessageBoxButton.OK,
                        MessageBoxImage.Information);
                }

                return;
            }

            IsConnected = false;
            SetDetachedModelInfo("Not yet attached");
            StatusText = string.IsNullOrWhiteSpace(result.Message)
                ? $"{_productName} connection unavailable."
                : result.Message;

            if (showMessage)
            {
                MessageBox.Show(
                    StatusText,
                    ProductTitle,
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
            }
        }

        private void CloseCurrentInstance()
        {
            var result = _closeCurrentInstanceUseCase.Execute();

            if (result.IsSuccess)
            {
                IsConnected = false;
                SetDetachedModelInfo("Not connected");
                StatusText = result.Message;

                MessageBox.Show(
                    result.Message,
                    ProductTitle,
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
                return;
            }

            MessageBox.Show(
                result.Message,
                ProductTitle,
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
            ShowOperationResult(_getSelectedCSISapModelPointsUseCase.Execute());
        }

        private void GetSelectedFrames()
        {
            ShowOperationResult(_getSelectedCSISapModelFramesUseCase.Execute());
        }

        private void CreateShellAreasFromSelectedFrames()
        {
            var propertyName = PromptForShellPropertyName();
            if (propertyName == null)
            {
                return;
            }

            ShowOperationResult(_createShellAreasFromSelectedFramesUseCase.Execute(propertyName));
        }

        private static string PromptForShellPropertyName()
        {
            var dialog = new Window
            {
                Title = "Shell Property",
                Width = 360,
                Height = 150,
                MinWidth = 360,
                MinHeight = 150,
                ResizeMode = ResizeMode.NoResize,
                WindowStartupLocation = WindowStartupLocation.CenterScreen,
                FontFamily = new System.Windows.Media.FontFamily("Segoe UI"),
                FontSize = 12
            };

            var root = new StackPanel { Margin = new Thickness(14) };
            var label = new TextBlock
            {
                Text = "Enter shell property name. Leave blank to use Default.",
                Margin = new Thickness(0, 0, 0, 8),
                TextWrapping = TextWrapping.Wrap
            };
            var textBox = new TextBox
            {
                Text = "Default",
                Height = 26,
                VerticalContentAlignment = VerticalAlignment.Center,
                Margin = new Thickness(0, 0, 0, 12)
            };

            var buttons = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right
            };
            var okButton = new Button
            {
                Content = "OK",
                Width = 74,
                Margin = new Thickness(0, 0, 8, 0),
                IsDefault = true
            };
            var cancelButton = new Button
            {
                Content = "Cancel",
                Width = 74,
                IsCancel = true
            };

            okButton.Click += delegate
            {
                dialog.DialogResult = true;
                dialog.Close();
            };
            cancelButton.Click += delegate
            {
                dialog.DialogResult = false;
                dialog.Close();
            };

            buttons.Children.Add(okButton);
            buttons.Children.Add(cancelButton);
            root.Children.Add(label);
            root.Children.Add(textBox);
            root.Children.Add(buttons);
            dialog.Content = root;

            var result = dialog.ShowDialog();
            return result == true ? textBox.Text : null;
        }

        private void ShowOperationResult(OperationResult result)
        {
            MessageBox.Show(
                result.Message,
                ProductTitle,
                MessageBoxButton.OK,
                result.IsSuccess ? MessageBoxImage.Information : MessageBoxImage.Warning);
        }

        private void GetLoadPatterns()
        {
            var result = _getLoadPatternsUseCase.Execute();
            if (result.IsSuccess)
            {
                LoadPatterns.Clear();
                if (result.Data != null)
                {
                    foreach (var p in result.Data)
                    {
                        LoadPatterns.Add(p);
                    }
                }
            }
            else
            {
                ShowOperationResult(OperationResult.Failure(result.Message));
            }
        }

        private void DeleteSelectedLoadPatterns(System.Collections.IList selectedItems)
        {
            if (selectedItems == null || selectedItems.Count == 0) return;
            
            var list = new System.Collections.Generic.List<string>();
            foreach (var item in selectedItems)
            {
                if (item is ExcelCSIToolBoxAddIn.Data.DTOs.CSISapModelLoadPatternDTO dto)
                {
                    list.Add(dto.Name);
                }
            }

            if (list.Count == 0) return;

            var result = _deleteLoadPatternsUseCase.Execute(list);
            ShowOperationResult(result);
            if (result.IsSuccess)
            {
                GetLoadPatterns(); // refresh list after deletion
            }
        }

        private void GetLoadCombinations()
        {
            var result = _getLoadCombinationsUseCase.Execute();
            if (result.IsSuccess)
            {
                LoadCombinations.Clear();
                if (result.Data != null)
                {
                    foreach (var c in result.Data)
                    {
                        LoadCombinations.Add(c);
                    }
                }
            }
            else
            {
                ShowOperationResult(OperationResult.Failure(result.Message));
            }
        }

        private void DeleteSelectedLoadCombinations(System.Collections.IList selectedItems)
        {
            if (selectedItems == null || selectedItems.Count == 0) return;
            
            var list = new System.Collections.Generic.List<string>();
            foreach (var item in selectedItems)
            {
                if (item is ExcelCSIToolBoxAddIn.Data.DTOs.CSISapModelLoadCombinationDTO dto)
                {
                    list.Add(dto.Name);
                }
            }

            if (list.Count == 0) return;

            var result = _deleteLoadCombinationsUseCase.Execute(list);
            ShowOperationResult(result);
            if (result.IsSuccess)
            {
                GetLoadCombinations(); // refresh list after deletion
            }
        }

        private void ViewLoadCombination(System.Collections.IList selectedItems)
        {
            if (selectedItems == null || selectedItems.Count == 0) return;
            
            // Only view the first selected item
            var firstItem = selectedItems[0];
            if (firstItem is ExcelCSIToolBoxAddIn.Data.DTOs.CSISapModelLoadCombinationDTO dto)
            {
                var result = _getLoadCombinationDetailsUseCase.Execute(dto.Name);
                if (result.IsSuccess)
                {
                    var window = new ExcelCSIToolBoxAddIn.UI.Views.LoadCombinationDetailsWindow(result.Data);
                    window.Owner = System.Windows.Application.Current.MainWindow;
                    window.ShowDialog();
                }
                else
                {
                    ShowOperationResult(OperationResult.Failure(result.Message));
                }
            }
        }

        private void ShowPlaceholder(string featureName)
        {
            MessageBox.Show(
                $"{featureName} is a placeholder for phase 1.",
                ProductTitle,
                MessageBoxButton.OK,
                MessageBoxImage.Information);
        }
    }
}
