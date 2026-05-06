using System.Collections.Specialized;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Media;
using ExcelCSIToolBoxAddIn.UI.Helpers;
using ExcelCSIToolBoxAddIn.UI.ViewModels;

namespace ExcelCSIToolBoxAddIn.UI.Views
{
    public partial class LoadCombinationMatrixView : Window
    {
        private LoadCombinationMatrixViewModel _viewModel;
        private readonly BlankFactorCellBackgroundConverter _factorBackgroundConverter = new BlankFactorCellBackgroundConverter();

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
            }

            _viewModel = viewModel;
            if (_viewModel != null)
            {
                _viewModel.RequestClose += ViewModel_RequestClose;
                _viewModel.LoadPatternNames.CollectionChanged += LoadPatternNames_CollectionChanged;
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
                Header = "LoadCombinationName",
                Binding = new Binding("LoadCombinationName")
                {
                    Mode = BindingMode.TwoWay,
                    UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged
                },
                MinWidth = 180,
                Width = new DataGridLength(260),
                ElementStyle = CreateTextStyle(TextAlignment.Left, Brushes.Blue, CreateSolidBrush(248, 226, 211)),
                EditingElementStyle = CreateTextBoxStyle(TextAlignment.Left, Brushes.Blue, CreateSolidBrush(248, 226, 211))
            });

            foreach (string patternName in _viewModel.LoadPatternNames)
            {
                var binding = new Binding("[" + patternName + "]")
                {
                    Mode = BindingMode.TwoWay,
                    UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged,
                    TargetNullValue = string.Empty
                };

                MatrixGrid.Columns.Add(new DataGridTextColumn
                {
                    Header = patternName,
                    Binding = binding,
                    MinWidth = 90,
                    Width = new DataGridLength(1, DataGridLengthUnitType.Auto),
                    ElementStyle = CreateFactorTextStyle(patternName),
                    EditingElementStyle = CreateFactorTextBoxStyle(patternName)
                });
            }
        }

        private static Style CreateTextStyle(TextAlignment alignment, Brush foreground, Brush background)
        {
            var style = new Style(typeof(TextBlock));
            style.Setters.Add(new Setter(TextBlock.TextAlignmentProperty, alignment));
            style.Setters.Add(new Setter(TextBlock.ForegroundProperty, foreground));
            style.Setters.Add(new Setter(TextBlock.BackgroundProperty, background));
            style.Setters.Add(new Setter(TextBlock.PaddingProperty, new Thickness(4, 1, 4, 1)));
            return style;
        }

        private static Style CreateTextBoxStyle(TextAlignment alignment, Brush foreground, Brush background)
        {
            var style = new Style(typeof(TextBox));
            style.Setters.Add(new Setter(TextBox.TextAlignmentProperty, alignment));
            style.Setters.Add(new Setter(TextBox.ForegroundProperty, foreground));
            style.Setters.Add(new Setter(TextBox.BackgroundProperty, background));
            style.Setters.Add(new Setter(TextBox.BorderThicknessProperty, new Thickness(0)));
            style.Setters.Add(new Setter(TextBox.PaddingProperty, new Thickness(4, 1, 4, 1)));
            return style;
        }

        private Style CreateFactorTextStyle(string patternName)
        {
            var style = CreateTextStyle(TextAlignment.Center, Brushes.Blue, Brushes.Transparent);
            style.Setters.Add(new Setter(TextBlock.BackgroundProperty, CreateBackgroundBinding(patternName)));
            return style;
        }

        private Style CreateFactorTextBoxStyle(string patternName)
        {
            var style = CreateTextBoxStyle(TextAlignment.Center, Brushes.Blue, Brushes.Transparent);
            style.Setters.Add(new Setter(TextBox.BackgroundProperty, CreateBackgroundBinding(patternName)));
            return style;
        }

        private Binding CreateBackgroundBinding(string patternName)
        {
            return new Binding("[" + patternName + "]")
            {
                Converter = _factorBackgroundConverter
            };
        }

        private static Brush CreateSolidBrush(byte r, byte g, byte b)
        {
            return new SolidColorBrush(Color.FromRgb(r, g, b));
        }
    }
}
