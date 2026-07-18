using System.Collections.Generic;

namespace ExcelCSIToolBox.Application.Modelling.DropPanels
{
    public sealed class DropPanelTableRecord
    {
        public DropPanelTableRecord()
        {
            Values = new Dictionary<string, string>(System.StringComparer.OrdinalIgnoreCase);
        }

        public Dictionary<string, string> Values { get; set; }
    }
}
