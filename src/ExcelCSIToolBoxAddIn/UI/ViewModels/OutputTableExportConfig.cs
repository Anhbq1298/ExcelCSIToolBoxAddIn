using System;

namespace ExcelCSIToolBoxAddIn.UI.ViewModels
{
    public sealed class OutputTableExportConfig
    {
        public OutputTableExportConfig()
        {
            PopupProfileKey = "ForceOutput";
        }

        public string TableDisplayName { get; set; }

        public string Breadcrumb { get; set; }

        public string Description { get; set; }

        public string PopupProfileKey { get; set; }

        public string DefaultSelectedCaseOrCombo { get; set; }

        public string WorksheetNamePrefix { get; set; }

        public string EmptyDataMessage { get; set; }

        public BaseReactionUnitOption ExportUnitOption { get; set; }

        public Func<bool, object[,]> StaticExportValuesFactory { get; set; }

        public int StaticRecordCount { get; set; }

        public string StaticSuccessMessage { get; set; }

        public bool DefaultAddHeaders { get; set; }

        public static OutputTableExportConfig ForTable(string tableDisplayName, string breadcrumb)
        {
            return new OutputTableExportConfig
            {
                TableDisplayName = string.IsNullOrWhiteSpace(tableDisplayName) ? "Base Reactions" : tableDisplayName,
                Breadcrumb = breadcrumb
            };
        }

        internal OutputTableExportConfig Normalize()
        {
            TableDisplayName = string.IsNullOrWhiteSpace(TableDisplayName)
                ? "Base Reactions"
                : TableDisplayName.Trim();

            if (string.IsNullOrWhiteSpace(Breadcrumb))
            {
                Breadcrumb = "ETABS Toolbox / ANALYSIS RESULTS / " + TableDisplayName;
            }

            if (string.IsNullOrWhiteSpace(WorksheetNamePrefix))
            {
                WorksheetNamePrefix = TableDisplayName;
            }

            if (string.IsNullOrWhiteSpace(PopupProfileKey))
            {
                PopupProfileKey = "ForceOutput";
            }

            if (string.IsNullOrWhiteSpace(EmptyDataMessage))
            {
                EmptyDataMessage = "ETABS returned no " + TableDisplayName + " records for the selected cases/combinations.";
            }

            return this;
        }
    }
}
