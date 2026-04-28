using System.Collections.Generic;

namespace ExcelCSIToolBox.Data.Models
{
    public sealed class CSISapModelShellCoordinateInput
    {
        public double X { get; set; }
        public double Y { get; set; }
        public double Z { get; set; }
    }

    public sealed class CSISapModelShellByCoordInput
    {
        public IReadOnlyList<CSISapModelShellCoordinateInput> Points { get; set; }
        public string PropertyName { get; set; }
        public string UserName { get; set; }
        public string CoordinateSystem { get; set; }
    }

    public sealed class CSISapModelShellByPointInput
    {
        public IReadOnlyList<string> PointNames { get; set; }
        public string PropertyName { get; set; }
        public string UserName { get; set; }
    }
}
