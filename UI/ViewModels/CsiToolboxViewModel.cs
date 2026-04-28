using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Globalization;
using ExcelCSIToolBoxAddIn.Common.Commands;
using ExcelCSIToolBoxAddIn.Common.Results;
using ExcelCSIToolBoxAddIn.Core.Application;
using ExcelCSIToolBoxAddIn.Data;
using ExcelCSIToolBoxAddIn.Data.DTOs;
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

        private readonly GetFrameSectionsUseCase _getFrameSectionsUseCase;
        private readonly GetFrameSectionDetailUseCase _getFrameSectionDetailUseCase;
        private readonly UpdateFrameSectionUseCase _updateFrameSectionUseCase;
        private readonly RenameFrameSectionUseCase _renameFrameSectionUseCase;

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

            _getFrameSectionsUseCase = new GetFrameSectionsUseCase(csiConnectionService);
            _getFrameSectionDetailUseCase = new GetFrameSectionDetailUseCase(csiConnectionService);
            _updateFrameSectionUseCase = new UpdateFrameSectionUseCase(csiConnectionService);
            _renameFrameSectionUseCase = new RenameFrameSectionUseCase(csiConnectionService);

            LoadCombinations = new System.Collections.ObjectModel.ObservableCollection<ExcelCSIToolBoxAddIn.Data.DTOs.CSISapModelLoadCombinationDTO>();
            LoadPatterns = new System.Collections.ObjectModel.ObservableCollection<ExcelCSIToolBoxAddIn.Data.DTOs.CSISapModelLoadPatternDTO>();
            FrameSections = new System.Collections.ObjectModel.ObservableCollection<CSISapModelFrameSectionDTO>();
            SectionDimensionAnnotations = new System.Collections.ObjectModel.ObservableCollection<SectionDimensionAnnotation>();

            AttachToRunningCsiCommand = new RelayCommand(() => LoadConnectionState(showMessage: true));
            CloseCurrentInstanceCommand = new RelayCommand(CloseCurrentInstance, () => IsConnected);

            CreateIshapeSectionCommand = new RelayCommand(() => ShowOperationResult(_createSteelISectionsUseCase.Execute()), () => IsConnected);
            CreateChannelSectionCommand = new RelayCommand(() => ShowOperationResult(_createSteelChannelSectionsUseCase.Execute()), () => IsConnected);
            CreateAngleSectionCommand = new RelayCommand(() => ShowOperationResult(_createSteelAngleSectionsUseCase.Execute()), () => IsConnected);
            CreateTubeSectionCommand = new RelayCommand(() => ShowOperationResult(_createSteelTubeSectionsUseCase.Execute()), () => IsConnected);
            CreatePipeSectionCommand = new RelayCommand(() => ShowOperationResult(_createSteelPipeSectionsUseCase.Execute()), () => IsConnected);

            CreateConcreteRectangleSectionCommand = new RelayCommand(() => ShowOperationResult(_createConcreteRectangleSectionsUseCase.Execute()), () => IsConnected);
            CreateConcreteCircleSectionCommand = new RelayCommand(() => ShowOperationResult(_createConcreteCircleSectionsUseCase.Execute()), () => IsConnected);

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
            AddLoadCombinationFromExcelCommand = new RelayCommand(() => ShowPlaceholder("Add Load Combination From Excel"), () => IsConnected);
            DeleteSelectedLoadCombinationsCommand = new RelayCommand<System.Collections.IList>(DeleteSelectedLoadCombinations, _ => IsConnected);
            ViewLoadCombinationCommand = new RelayCommand<System.Collections.IList>(ViewLoadCombination, _ => IsConnected);
            
            GetFrameSectionsCommand = new RelayCommand(GetFrameSections, () => IsConnected);
            EditFrameSectionCommand = new RelayCommand<CSISapModelFrameSectionDTO>(EditFrameSection, _ => IsConnected);

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
                System.Windows.Input.CommandManager.InvalidateRequerySuggested();
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
                RefreshSectionDimensionAnnotations();
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
        
        public ICommand GetFrameSectionsCommand { get; }
        public ICommand EditFrameSectionCommand { get; }

        public System.Collections.ObjectModel.ObservableCollection<ExcelCSIToolBoxAddIn.Data.DTOs.CSISapModelLoadPatternDTO> LoadPatterns { get; }
        public System.Collections.ObjectModel.ObservableCollection<ExcelCSIToolBoxAddIn.Data.DTOs.CSISapModelLoadCombinationDTO> LoadCombinations { get; }
        public System.Collections.ObjectModel.ObservableCollection<CSISapModelFrameSectionDTO> FrameSections { get; }
        public System.Collections.ObjectModel.ObservableCollection<SectionDimensionAnnotation> SectionDimensionAnnotations { get; }
        
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

        private void RefreshSectionDimensionAnnotations()
        {
            SectionDimensionAnnotations.Clear();

            if (SelectedFrameSectionDetail == null)
            {
                return;
            }

            foreach (var annotation in CreateDimensionAnnotations(SelectedFrameSectionDetail, GetLengthUnitText()))
            {
                SectionDimensionAnnotations.Add(annotation);
            }
        }

        private static System.Collections.Generic.IEnumerable<SectionDimensionAnnotation> CreateDimensionAnnotations(CSISapModelFrameSectionDetailDTO detail, string unit)
        {
            switch (detail.ShapeType)
            {
                case FrameSectionShapeType.Rectangular:
                    foreach (var item in DimensionItems(detail, unit,
                        Spec("h", "Depth ( t3 )", "Total depth ( t3 )"),
                        Spec("b", "Width ( t2 )", "Flange width ( t2 )")))
                    {
                        yield return item;
                    }
                    break;

                case FrameSectionShapeType.Tube:
                    foreach (var item in DimensionItems(detail, unit,
                        Spec("h", "Total depth ( t3 )", "Depth ( t3 )"),
                        Spec("b", "Flange width ( t2 )", "Width ( t2 )"),
                        Spec("t2", "Flange thickness ( tf )"),
                        Spec("t3", "Web thickness ( tw )")))
                    {
                        yield return item;
                    }
                    break;

                case FrameSectionShapeType.I:
                    foreach (var item in DimensionItems(detail, unit,
                        Spec("h", "Total depth ( t3 )", "Depth ( t3 )"),
                        Spec("b", "Top flange width ( t2 )", "Flange width ( t2 )"),
                        Spec("tw", "Web thickness ( tw )"),
                        Spec("tf", "Top flange thickness ( tf )", "Flange thickness ( tf )"),
                        Spec("t2b", "Bottom flange width ( t2b )"),
                        Spec("tfb", "Bottom flange thickness ( tfb )")))
                    {
                        yield return item;
                    }
                    break;

                case FrameSectionShapeType.Channel:
                case FrameSectionShapeType.Angle:
                case FrameSectionShapeType.DoubleAngle:
                    foreach (var item in DimensionItems(detail, unit,
                        Spec("h", "Total depth ( t3 )", "Depth ( t3 )"),
                        Spec("b", "Flange width ( t2 )", "Width ( t2 )"),
                        Spec("tw", "Web thickness ( tw )"),
                        Spec("tf", "Flange thickness ( tf )"),
                        Spec("dis", "Spacing ( dis )")))
                    {
                        yield return item;
                    }
                    break;

                case FrameSectionShapeType.Pipe:
                    foreach (var item in DimensionItems(detail, unit,
                        Spec("d", "Outside diameter ( t3 )", "Diameter ( t3 )"),
                        Spec("t", "Wall thickness ( tw )", "Wall thickness ( t )")))
                    {
                        yield return item;
                    }
                    break;

                case FrameSectionShapeType.Circular:
                    foreach (var item in DimensionItems(detail, unit,
                        Spec("d", "Diameter ( t3 )", "Outside diameter ( t3 )"),
                        Spec("r", "Radius ( r )")))
                    {
                        yield return item;
                    }
                    break;

                case FrameSectionShapeType.General:
                    foreach (var item in DimensionItems(detail, unit,
                        Spec("h", "Total depth ( t3 )", "Depth ( t3 )"),
                        Spec("b", "Width ( t2 )")))
                    {
                        yield return item;
                    }
                    break;
            }
        }

        private static System.Collections.Generic.IEnumerable<SectionDimensionAnnotation> DimensionItems(
            CSISapModelFrameSectionDetailDTO detail,
            string unit,
            params DimensionSpec[] specs)
        {
            foreach (var spec in specs)
            {
                if (TryGetDimensionValue(detail, out double value, spec.DimensionNames))
                {
                    yield return CreateDimensionItem(spec.Key, value, unit, detail.ShapeType.ToString());
                }
            }
        }

        private static SectionDimensionAnnotation CreateDimensionItem(string key, double value, string unit, string sectionType)
        {
            string valueText = value.ToString("0.###", CultureInfo.InvariantCulture);
            string displayText = string.IsNullOrWhiteSpace(unit)
                ? $"{key} = {valueText}"
                : $"{key} = {valueText} {unit}";

            return new SectionDimensionAnnotation
            {
                Key = key,
                DisplayLabel = key,
                Value = value,
                Unit = unit,
                DisplayText = displayText,
                DescriptionText = $"{key} = {GetDimensionDescription(key, sectionType)}",
                SectionType = sectionType
            };
        }

        private static string GetDimensionDescription(string key, string sectionType)
        {
            switch (key)
            {
                case "h": return "height";
                case "b": return "width";
                case "d": return "diameter";
                case "r": return "radius";
                case "t": return "thickness";
                case "tw": return "web thickness";
                case "tf": return "flange thickness";
                case "t2": return sectionType == FrameSectionShapeType.Tube.ToString() ? "top/bottom wall thickness" : "local 2 dimension";
                case "t3": return sectionType == FrameSectionShapeType.Tube.ToString() ? "side wall thickness" : "local 3 dimension";
                case "t2b": return "bottom flange width";
                case "tfb": return "bottom flange thickness";
                case "dis": return "spacing";
                default: return "dimension";
            }
        }

        private static DimensionSpec Spec(string key, params string[] dimensionNames)
        {
            return new DimensionSpec { Key = key, DimensionNames = dimensionNames };
        }

        private static bool TryGetDimensionValue(CSISapModelFrameSectionDetailDTO detail, out double value, params string[] keys)
        {
            value = 0;
            if (detail?.Dimensions == null)
            {
                return false;
            }

            foreach (string key in keys)
            {
                if (detail.Dimensions.TryGetValue(key, out value))
                {
                    return true;
                }
            }

            return false;
        }

        private string GetLengthUnitText()
        {
            string unitText = CurrentModelUnitText ?? string.Empty;
            string lower = unitText.ToLowerInvariant();

            if (lower.Contains("mm")) return "mm";
            if (lower.Contains("cm")) return "cm";
            if (lower.Contains("-m-") || lower.EndsWith("-m")) return "m";
            if (lower.Contains("in")) return "in";
            if (lower.Contains("ft")) return "ft";

            return string.Empty;
        }

        private class DimensionSpec
        {
            public string Key { get; set; }
            public string[] DimensionNames { get; set; }
        }

        private void LoadSelectedSectionDetail(CSISapModelFrameSectionDTO section)
        {
            if (section == null || _getFrameSectionDetailUseCase == null)
            {
                SelectedFrameSectionDetail = null;
                return;
            }
            var result = _getFrameSectionDetailUseCase.Execute(section.Name);
            SelectedFrameSectionDetail = result.IsSuccess ? result.Data : null;
        }

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
                StatusText = "Attached successfully.";
                
                // Automatically refresh lists when connection is established
                GetLoadPatterns();
                GetLoadCombinations();
                GetFrameSections();

                if (showMessage)
                {
                    ShowOperationResult(OperationResult.Success("Successfully attached to the running application."));
                }

                return;
            }

            IsConnected = false;
            SetDetachedModelInfo("Not yet attached");
            StatusText = string.IsNullOrWhiteSpace(result.Message)
                ? $"{_productName} connection unavailable."
                : result.Message;
                
            LoadPatterns.Clear();
            LoadCombinations.Clear();
            FrameSections.Clear();
            SelectedFrameSection = null;

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

                LoadPatterns.Clear();
                LoadCombinations.Clear();
                FrameSections.Clear();
                SelectedFrameSection = null;

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
                    window.ShowDialog();
                }
                else
                {
                    ShowOperationResult(OperationResult.Failure(result.Message));
                }
            }
        }

        private void GetFrameSections()
        {
            var result = _getFrameSectionsUseCase.Execute();
            if (result.IsSuccess)
            {
                FrameSections.Clear();
                if (result.Data != null)
                {
                    foreach (var section in result.Data)
                    {
                        FrameSections.Add(section);
                    }
                }
            }
            else
            {
                ShowOperationResult(OperationResult.Failure(result.Message));
            }
        }

        private void EditFrameSection(CSISapModelFrameSectionDTO section)
        {
            if (section == null) return;

            var result = _getFrameSectionDetailUseCase.Execute(section.Name);
            if (result.IsSuccess)
            {
                var window = new ExcelCSIToolBoxAddIn.UI.Views.FrameSectionDetailWindow(result.Data);
                bool? dialogResult = window.ShowDialog();
                if (dialogResult != true)
                {
                    return;
                }

                OperationResult saveResult;
                string selectedName;
                if (window.ViewModel.IsRename)
                {
                    var confirm = MessageBox.Show(
                        "Renaming a section will create a new section, reassign frames using the old section, and then delete the old section when possible. Continue?",
                        ProductTitle,
                        MessageBoxButton.OKCancel,
                        MessageBoxImage.Warning);

                    if (confirm != MessageBoxResult.OK)
                    {
                        return;
                    }

                    var renameInput = window.ViewModel.ToRenameDto();
                    selectedName = renameInput.SectionName;
                    saveResult = _renameFrameSectionUseCase.Execute(renameInput);
                }
                else
                {
                    var updateInput = window.ViewModel.ToUpdateDto();
                    selectedName = updateInput.SectionName;
                    saveResult = _updateFrameSectionUseCase.Execute(updateInput);
                }

                ShowOperationResult(saveResult);
                if (saveResult.IsSuccess)
                {
                    GetFrameSections();
                    SelectFrameSectionByName(selectedName);
                }
            }
            else
            {
                ShowOperationResult(OperationResult.Failure(result.Message));
            }
        }

        private void SelectFrameSectionByName(string sectionName)
        {
            foreach (var section in FrameSections)
            {
                if (string.Equals(section.Name, sectionName, System.StringComparison.Ordinal))
                {
                    SelectedFrameSection = section;
                    return;
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
