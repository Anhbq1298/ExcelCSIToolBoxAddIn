using System.Collections.Generic;

namespace ExcelCSIToolBox.Data.DTOs.CSI
{
    public class CSISapModelFrameSectionUpdateDTO
    {
        public string OriginalName { get; set; }
        public string SectionName { get; set; }
        public string MaterialName { get; set; }
        public FrameSectionShapeType ShapeType { get; set; }
        public Dictionary<string, double> Dimensions { get; set; } = new Dictionary<string, double>();
        public int Color { get; set; }
        public string Notes { get; set; }
    }

    public class CSISapModelFrameSectionRenameDTO : CSISapModelFrameSectionUpdateDTO
    {
    }
}

