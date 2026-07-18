using System.Collections.Generic;

namespace ExcelCSIToolBox.Application.Modelling.DropPanels
{
    public sealed class DropPanelApplyResult
    {
        public DropPanelApplyResult()
        {
            CreatedAreaNames = new List<string>();
            VerificationIssues = new List<DropPanelVerificationIssue>();
            LogEntries = new List<DropPanelLogEntry>();
        }

        public string BackupFilePath { get; set; }

        public List<string> CreatedAreaNames { get; set; }

        public List<DropPanelVerificationIssue> VerificationIssues { get; set; }

        public List<DropPanelLogEntry> LogEntries { get; set; }

        public bool VerificationPassed
        {
            get { return VerificationIssues.Count == 0; }
        }
    }
}
