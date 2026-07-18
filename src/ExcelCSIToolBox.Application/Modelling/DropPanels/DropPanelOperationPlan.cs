using System.Collections.Generic;

namespace ExcelCSIToolBox.Application.Modelling.DropPanels
{
    public sealed class DropPanelOperationPlan
    {
        public DropPanelOperationPlan()
        {
            Columns = new List<DropPanelColumnInfo>();
            SourceAreas = new List<DropPanelAreaInfo>();
            Openings = new List<DropPanelAreaInfo>();
            Requests = new List<DropPanelRequest>();
            Regions = new List<DropPanelRegion>();
            ValidationMessages = new List<string>();
        }

        public List<DropPanelColumnInfo> Columns { get; set; }

        public List<DropPanelAreaInfo> SourceAreas { get; set; }

        public List<DropPanelAreaInfo> Openings { get; set; }

        public List<DropPanelRequest> Requests { get; set; }

        public List<DropPanelRegion> Regions { get; set; }

        public List<string> ValidationMessages { get; set; }

        public string ModelPath { get; set; }

        public string PresentUnits { get; set; }

        public bool IsValid
        {
            get { return ValidationMessages.Count == 0 && Regions.Count > 0; }
        }
    }
}
