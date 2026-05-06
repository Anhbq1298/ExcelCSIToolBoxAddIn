using System;
using System.Globalization;
using System.Windows.Data;
using System.Windows.Media;

namespace ExcelCSIToolBoxAddIn.UI.Helpers
{
    public class BlankFactorCellBackgroundConverter : IValueConverter
    {
        private static readonly Brush BlankBrush = new SolidColorBrush(Color.FromRgb(210, 210, 210));
        private static readonly Brush ValueBrush = new SolidColorBrush(Color.FromRgb(248, 226, 211));

        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            string text = value == null ? null : value.ToString();
            return string.IsNullOrWhiteSpace(text) ? BlankBrush : ValueBrush;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return Binding.DoNothing;
        }
    }
}
