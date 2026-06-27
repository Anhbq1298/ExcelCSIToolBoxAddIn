namespace ExcelCSIToolBoxAddIn.UI.ViewModels
{
    public sealed class OutputTablePopupProfile
    {
        public OutputTablePopupProfile()
        {
            Key = "ForceOutput";
            CaseSelectionMode = OutputCaseSelectionMode.AllCasesAndCombos;
            CaseSelectorTitle = "Load Case / Combination";
            AllowMultipleCases = true;
            ShowCaseComboSelector = true;
            ShowUnitSelector = true;
            ShowComboSelector = true;
            DefaultToCurrentEtabsUnit = true;
            EmptyDataMessage = "No records found.";
            WorksheetNamePrefix = "Output";
        }

        public string Key { get; set; }

        public OutputCaseSelectionMode CaseSelectionMode { get; set; }

        public string CaseSelectorTitle { get; set; }

        public bool AllowMultipleCases { get; set; }

        public bool ShowCaseComboSelector { get; set; }

        public bool ShowUnitSelector { get; set; }

        public bool ShowComboSelector { get; set; }

        public bool DefaultToCurrentEtabsUnit { get; set; }

        public string EmptyDataMessage { get; set; }

        public string WorksheetNamePrefix { get; set; }
    }
}
