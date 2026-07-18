namespace ExcelCSIToolBox.Application.Modelling.DropPanels
{
    public sealed class DropPanelVerificationIssue
    {
        public string SourceAreaName { get; set; }

        public string NewAreaName { get; set; }

        public string AssignmentType { get; set; }

        public string ExpectedValue { get; set; }

        public string ActualValue { get; set; }

        public string ErrorMessage { get; set; }
    }
}
