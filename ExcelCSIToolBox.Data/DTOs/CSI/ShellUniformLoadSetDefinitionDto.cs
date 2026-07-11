using System.Collections.Generic;

namespace ExcelCSIToolBox.Data.DTOs.CSI
{
    public class ShellUniformLoadSetDefinitionDto
    {
        public string Name { get; set; }

        public Dictionary<string, double> LoadValuesByPattern { get; set; } = new Dictionary<string, double>();
    }
}
