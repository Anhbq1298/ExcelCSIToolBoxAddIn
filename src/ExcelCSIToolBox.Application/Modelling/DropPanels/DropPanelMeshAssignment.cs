using System.Collections.Generic;

namespace ExcelCSIToolBox.Application.Modelling.DropPanels
{
    public sealed class DropPanelMeshAssignment
    {
        public DropPanelMeshAssignment()
        {
            FieldKeys = new List<string>();
            Records = new List<DropPanelTableRecord>();
        }

        public string TableKey { get; set; }

        public int TableVersion { get; set; }

        public List<string> FieldKeys { get; set; }

        public List<DropPanelTableRecord> Records { get; set; }
    }
}
