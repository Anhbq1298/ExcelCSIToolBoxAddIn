using System.Collections.Generic;

namespace ExcelCSIToolBox.Application.Modelling.DropPanels
{
    public sealed class DropPanelPreparationSnapshot
    {
        public DropPanelPreparationSnapshot()
        {
            Areas = new List<DropPanelAreaInfo>();
            Openings = new List<DropPanelAreaInfo>();
        }

        public List<DropPanelAreaInfo> Areas { get; set; }

        public List<DropPanelAreaInfo> Openings { get; set; }

        public string ModelPath { get; set; }

        public string PresentUnits { get; set; }
    }
}
