namespace ExcelCSIToolBox.Core.Contracts.CSI
{
    public class CSISapModelRunningInstanceDTO
    {
        public string InstanceId { get; set; }
        public int? ProcessId { get; set; }
        public string DisplayName { get; set; }
        public string ModelPath { get; set; }
        public string ModelFileName { get; set; }
        public string ModelCurrentUnit { get; set; }
    }
}
