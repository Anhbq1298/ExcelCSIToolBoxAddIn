using System.Collections.Generic;

namespace ExcelCSIToolBoxAddIn.Data.DTOs
{
    public enum FrameSectionShapeType
    {
        Unknown, I, Channel, T, Angle, DoubleAngle, Tube, Pipe, Rectangular, Circular, General
    }

    public class CSISapModelFrameSectionDTO
    {
        public string Name { get; set; }
        public FrameSectionShapeType ShapeType { get; set; }
        public string MaterialName { get; set; }
    }

    public class CSISapModelFrameSectionDetailDTO : CSISapModelFrameSectionDTO
    {
        public Dictionary<string, double> Dimensions { get; set; } = new Dictionary<string, double>();
        public int Color { get; set; }
        public string Notes { get; set; }
    }
}
