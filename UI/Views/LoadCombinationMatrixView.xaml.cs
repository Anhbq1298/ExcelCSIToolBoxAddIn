using System.Collections.Specialized;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Media;
using ExcelCSIToolBoxAddIn.UI.ViewModels;

namespace ExcelCSIToolBoxAddIn.UI.Views
{
    public partial class LoadCombinationMatrixView : Window
    {
        private static readonly Brush TextBrush = CreateSolidBrush(47, 51, 55);
        private static readonly Brush WhiteBrush = Brushes.White;
        private static readonly Brush LightBorderBrush = CreateSolidBrush(218, 221, 226);
        private LoadCombinationMatrixViewModel _viewModel;

        public LoadCombinationMatrixView(LoadCombinationMatrixViewModel viewModel)
        {
            InitializeComponent();
            DataContext = viewModel;
            AttachViewModel(viewModel);
            RebuildColumns();
        }

        private void AttachViewModel(LoadCombinationMatrixViewModel viewModel)
        {
            if (_viewModel != null)
            {
                _viewModel.RequestClose -= ViewModel_RequestClose;
                _viewModel.LoadPatternNames.CollectionChanged -= LoadPatternNames_CollectionChanged;
                _viewModel.LoadCombinationReferenceNames.CollectionChanged -= LoadPatternNames_CollectionChanged;
            }

            _viewModel = viewModel;
            if (_viewModel != null)
            {
                _viewModel.RequestClose += ViewModel_RequestClose;
                _viewModel.LoadPatternNames.CollectionChanged += LoadPatternNames_CollectionChanged;
                _viewModel.LoadCombinationReferenceNames.CollectionChanged += LoadPatternNames_CollectionChanged;
            }
        }

        private void ViewModel_RequestClose(object sender, System.EventArgs e)
        {
            Close();
        }

        private void LoadPatternNames_CollectionChanged(object sender, NotifyCollectionChangedEventArgs e)
        {
            RebuildColumns();
        }

        private void RebuildColumns()
        {
            if (MatrixGrid == null || _viewModel == null)
            {
                return;
            }

            MatrixGrid.Columns.Clear();
            MatrixGrid.Columns.Add(new DataGridTextColumn
            {
                Header = "Load Combination Name",
                Binding = new Binding("LoadCombinationName")
                {
                    Mode = BindingMode.TwoWay,
                    UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged
                },
                MinWidth = 180,
                Width = new DataGridLength(260),
                HeaderStyle = FindHeaderStyle("NeutralMatrixHeaderStyle"),
                ElementStyle = CreateTextStyle(TextAlignment.Left),
                EditingElementStyle = CreateTextBoxStyle(TextAlignment.Left)
            });

            MatrixGrid.Columns.Add(new DataGridComboBoxColumn
            {
                Header = "Combination Type",
                ItemsSource = _viewModel.CombinationTypeOptions,
                SelectedValueBinding = new Binding("CombinationType")
                {
                    Mode = BindingMode.TwoWay,
                    UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged
                },
                SelectedValuePath = "Value",
                DisplayMemberPath = "DisplayName",
                MinWidth = 150,
                Width = new DataGridLength(170),
                HeaderStyle = FindHeaderStyle("NeutralMatrixHeaderStyle"),
                ElementStyle = CreateComboBoxStyle(false),
                EditingElementStyle = CreateComboBoxStyle(true)
            });

            foreach (string patternName in _viewModel.LoadPatternNames)
            {
                var binding = new Binding("LoadCaseFactors[" + patternName + "]")
                {
                    Mode = BindingMode.TwoWay,
                    UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged,
                    TargetNullValue = string.Empty
                };

                MatrixGrid.Columns.Add(new DataGridTextColumn
                {
                    Header = "LC | " + patternName,
                    Binding = binding,
                    MinWidth = 90,
                    Width = new DataGridLength(1, DataGridLengthUnitType.Auto),
                    HeaderStyle = FindHeaderStyle("NeutralMatrixHeaderStyle"),
                    CanUserReorder = false,
                    ElementStyle = CreateTextStyle(TextAlignment.Center),
                    EditingElementStyle = CreateTextBoxStyle(TextAlignment.Center)
                });
            }

            foreach (string comboName in _viewModel.LoadCombinationReferenceNames)
            {
                var binding = new Binding("LoadCombinationFactors[" + comboName + "]")
                {
                    Mode = BindingMode.TwoWay,
                    UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged,
                    TargetNullValue = string.Empty
                };

                MatrixGrid.Columns.Add(new DataGridTextColumn
                {
                    Header = "COMBO | " + comboName,
                    Binding = binding,
                    MinWidth = 100,
                    Width = new DataGridLength(1, DataGridLengthUnitType.Auto),
                    HeaderStyle = FindHeaderStyle("LoadCombinationHeaderStyle"),
                    CanUserReorder = false,
                    ElementStyle = CreateTextStyle(TextAlignment.Center),
                    EditingElementStyle = CreateTextBoxStyle(TextAlignment.Center)
                });
            }
        }

        private Style FindHeaderStyle(string resourceKey)
        {
            return MatrixGrid.TryFindResource(resourceKey) as Style;
        }

        private static Style CreateTextStyle(TextAlignment alignment)
        {
            var style = new Style(typeof(TextBlock));
            style.Setters.Add(new Setter(TextBlock.TextAlignmentProperty, alignment));
            style.Setters.Add(new Setter(TextBlock.ForegroundProperty, TextBrush));
            style.Setters.Add(new Setter(TextBlock.PaddingProperty, new Thickness(6, 2, 6, 2)));
            style.Setters.Add(new Setter(TextBlock.VerticalAlignmentProperty, VerticalAlignment.Center));
            return style;
        }

        private static Style CreateTextBoxStyle(TextAlignment alignment)
        {
            var style = new Style(typeof(TextBox));
            style.Setters.Add(new Setter(TextBox.TextAlignmentProperty, alignment));
            style.Setters.Add(new Setter(TextBox.ForegroundProperty, TextBrush));
            style.Setters.Add(new Setter(TextBox.BackgroundProperty, WhiteBrush));
            style.Setters.Add(new Setter(TextBox.BorderThicknessProperty, new Thickness(0)));
            style.Setters.Add(new Setter(TextBox.PaddingProperty, new Thickness(6, 2, 6, 2)));
            style.Setters.Add(new Setter(TextBox.VerticalContentAlignmentProperty, VerticalAlignment.Center));
            return style;
        }

        private static Style CreateComboBoxStyle(bool isEditing)
        {
            var style = new Style(typeof(ComboBox));
            style.Setters.Add(new Setter(Control.ForegroundProperty, TextBrush));
            style.Setters.Add(new Setter(Control.BackgroundProperty, WhiteBrush));
            style.Setters.Add(new Setter(Control.BorderBrushProperty, LightBorderBrush));
            style.Setters.Add(new Setter(Control.BorderThicknessProperty, isEditing ? new Thickness(1) : new Thickness(0)));
            style.Setters.Add(new Setter(Control.PaddingProperty, new Thickness(4, 1, 4, 1)));
            style.Setters.Add(new Setter(ComboBox.VerticalContentAlignmentProperty, VerticalAlignment.Center));
            style.Setters.Add(new Setter(ComboBox.IsSynchronizedWithCurrentItemProperty, false));
            if (!isEditing)
            {
                style.Setters.Add(new Setter(UIElement.IsHitTestVisibleProperty, false));
                style.Setters.Add(new Setter(Control.FocusableProperty, false));
            }

            return style;
        }

        private static Brush CreateSolidBrush(byte r, byte g, byte b)
        {
            return new SolidColorBrush(Color.FromRgb(r, g, b));
        }
    }
}
