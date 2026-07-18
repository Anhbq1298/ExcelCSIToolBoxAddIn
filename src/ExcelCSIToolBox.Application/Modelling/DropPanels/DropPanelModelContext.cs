namespace ExcelCSIToolBox.Application.Modelling.DropPanels
{
    public sealed class DropPanelModelContext
    {
        public string Version { get; set; }

        public string ModelFileName { get; set; }

        public string ModelPath { get; set; }

        public string PresentUnits { get; set; }

        public string LengthUnit { get; set; }

        public bool IsLocked { get; set; }
    }
}
