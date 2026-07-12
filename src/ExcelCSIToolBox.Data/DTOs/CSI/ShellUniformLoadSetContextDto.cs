using System.Collections.Generic;

namespace ExcelCSIToolBox.Data.DTOs.CSI
{
    public class ShellUniformLoadSetContextDto
    {
        public string ModelPath { get; set; }

        public string ModelFileName { get; set; }

        public string PresentUnitsText { get; set; }

        public List<string> LoadPatternNames { get; set; } = new List<string>();
    }
}
