using System.Collections.Generic;

namespace ExcelCSIToolBox.Application.Modelling.DropPanels
{
    public sealed class DropPanelApplyResult
    {
        public DropPanelApplyResult()
        {
            CreatedAreaNames = new List<string>();
            LogEntries = new List<DropPanelLogEntry>();
        }

        public List<string> CreatedAreaNames { get; set; }

        public List<DropPanelLogEntry> LogEntries { get; set; }

        public int ProcessedColumnCount { get; set; }

        public int CreatedDropAreaCount { get; set; }

        public string DropPropertyName { get; set; }

        public bool DropPropertyCreated { get; set; }

        public double DropThickness { get; set; }

        public string LengthUnit { get; set; }

        public string MaterialName { get; set; }
    }
}
