using System.Windows;

namespace ExcelCSIToolBoxAddIn.UI.ViewModels
{
    public class SectionDimensionAnnotation
    {
        public string Key { get; set; }
        public string DisplayLabel { get; set; }
        public string SectionType { get; set; }
        public Point StartPoint { get; set; }
        public Point EndPoint { get; set; }
        public string Orientation { get; set; }
        public Point LabelPosition { get; set; }
        public Visibility Visibility { get; set; } = Visibility.Visible;
    }
}
