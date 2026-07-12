using System.Collections.Generic;

namespace ExcelCSIToolBox.Data.DTOs.CSI
{
    public sealed class CSISapModelShellObjectDTO
    {
        public string Name { get; set; }
        public IReadOnlyList<string> PointNames { get; set; }
        public string PropertyName { get; set; }
        public bool IsSelected { get; set; }
    }

    public sealed class CSISapModelShellLoadDTO
    {
        public string AreaName { get; set; }
        public string LoadPattern { get; set; }
        public string LoadType { get; set; }
        public int Direction { get; set; }
        public double Value { get; set; }
        public string CoordinateSystem { get; set; }
    }
}
