using System.Collections.Generic;

namespace ExcelCSIToolBox.Application.Modelling.DropPanels
{
    public sealed class DropPanelRegion
    {
        public DropPanelRegion()
        {
            Points = new List<DropPanelPoint3D>();
            ColumnNames = new List<string>();
        }

        public string SourceAreaName { get; set; }

        public bool IsDrop { get; set; }

        public string ResultingSectionProperty { get; set; }

        public List<DropPanelPoint3D> Points { get; set; }

        public List<string> ColumnNames { get; set; }

        public DropPanelAreaAssignmentBackup Assignment { get; set; }
    }
}
