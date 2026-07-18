namespace ExcelCSIToolBox.Application.Modelling.DropPanels
{
    public sealed class DropPanelLogEntry
    {
        public string Timestamp { get; set; }

        public string EtabsModel { get; set; }

        public string Story { get; set; }

        public string Column { get; set; }

        public string SourceArea { get; set; }

        public string NewArea { get; set; }

        public string RegionType { get; set; }

        public string OriginalProperty { get; set; }

        public string NewProperty { get; set; }

        public string DirectLoadStatus { get; set; }

        public string ShellLoadSetStatus { get; set; }

        public string LocalAxisStatus { get; set; }

        public string Local3Status { get; set; }

        public string DiaphragmStatus { get; set; }

        public string Message { get; set; }
    }
}
