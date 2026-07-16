using System.Collections.Generic;

namespace ExcelCSIToolBox.Core.Contracts.CSI
{
    public class ShellUniformLoadSetAreaAssignmentDto
    {
        public string Story { get; set; }

        public string Label { get; set; }

        public string UniqueName { get; set; }

        public string LoadSetName { get; set; }
    }

    public class ShellUniformLoadSetSelectionResultDto
    {
        public int RequestedLoadSetCount { get; set; }

        public int MatchedLoadSetCount { get; set; }

        public int SelectedShellCount { get; set; }

        public int UnresolvedAreaCount { get; set; }

        public int DuplicateShellCount { get; set; }

        public int UnknownLoadSetCount { get; set; }

        public string SelectedStoryName { get; set; }

        public List<string> SelectedStoryNames { get; set; } = new List<string>();

        public string Message { get; set; }

        public string WarningMessage { get; set; }

        public List<string> SelectedShellNames { get; set; } = new List<string>();

        public List<string> UnknownLoadSetNames { get; set; } = new List<string>();

        public List<string> UnresolvedAreaReferences { get; set; } = new List<string>();
    }
}
