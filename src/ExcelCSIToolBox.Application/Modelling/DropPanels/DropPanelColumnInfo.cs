namespace ExcelCSIToolBox.Application.Modelling.DropPanels
{
    public sealed class DropPanelColumnInfo
    {
        public string FrameName { get; set; }

        public string BottomPointName { get; set; }

        public string TopPointName { get; set; }

        public string StoryName { get; set; }

        public double X { get; set; }

        public double Y { get; set; }

        public double Z { get; set; }

        public string SectionProperty { get; set; }

        public double LocalAxisRotationDegrees { get; set; }

        public bool IsValid { get; set; }

        public string ValidationStatus
        {
            get { return IsValid ? "Valid" : "Invalid"; }
        }

        public string ValidationMessage { get; set; }
    }
}
