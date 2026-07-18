using System.Collections.Generic;

namespace ExcelCSIToolBox.Application.Modelling.DropPanels
{
    public sealed class DropPanelAreaInfo
    {
        public DropPanelAreaInfo()
        {
            Points = new List<DropPanelPoint3D>();
            ConnectedColumnNames = new List<string>();
        }

        public string AreaName { get; set; }

        public string StoryName { get; set; }

        public string SectionProperty { get; set; }

        public double Elevation { get; set; }

        public bool IsOpening { get; set; }

        public List<DropPanelPoint3D> Points { get; set; }

        public List<string> ConnectedColumnNames { get; set; }

        public DropPanelAreaAssignmentBackup Assignment { get; set; }
    }
}
