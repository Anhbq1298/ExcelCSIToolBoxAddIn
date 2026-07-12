namespace ExcelCSIToolBoxAddIn.UI.ViewModels
{
    public class CsiRunningInstanceViewModel
    {
        public string InstanceId { get; set; }
        public int? ProcessId { get; set; }
        public string DisplayName { get; set; }
        public string ModelPath { get; set; }
        public string ModelFileName { get; set; }
        public string ModelCurrentUnit { get; set; }

        public string ToolTipText
        {
            get
            {
                string processText = ProcessId.HasValue ? $"Process ID: {ProcessId.Value}" : "Process ID: unknown";
                return string.IsNullOrWhiteSpace(ModelPath)
                    ? processText
                    : $"{processText}\n{ModelPath}";
            }
        }
    }
}
