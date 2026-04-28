namespace ExcelCSIToolBoxAddIn.Infrastructure.Csi
{
    /// <summary>
    /// Value object containing CSI connection details needed by the UI.
    /// </summary>
    public class CsiConnectionInfo
    {
        public bool IsConnected { get; set; }

        public string ModelPath { get; set; }

        public string ModelFileName { get; set; }

        public string ModelCurrentUnit { get; set; }

        /// <summary>
        /// Optional COM object references for future CSI operations.
        /// </summary>
        public object CsiObject { get; set; }

        public object SapModel { get; set; }
    }
}
