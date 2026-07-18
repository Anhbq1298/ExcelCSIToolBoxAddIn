using System.Collections.Generic;

namespace ExcelCSIToolBox.Application.Modelling.DropPanels
{
    public sealed class DropPanelRequest
    {
        public DropPanelRequest()
        {
            Points = new List<DropPanelPoint3D>();
        }

        public string ColumnName { get; set; }

        public string StoryName { get; set; }

        public double Elevation { get; set; }

        public double RotationDegrees { get; set; }

        public string DropProperty { get; set; }

        public List<DropPanelPoint3D> Points { get; set; }
    }
}
