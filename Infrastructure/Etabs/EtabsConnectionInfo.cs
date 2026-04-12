namespace ExcelCSIToolBoxAddIn.Infrastructure.Etabs
{
    /// <summary>
    /// Value object containing ETABS connection details needed by the UI.
    /// </summary>
    public class EtabsConnectionInfo
    {
        public bool IsConnected { get; set; }

        public string ModelFileName { get; set; }

        /// <summary>
        /// Optional COM object references for future ETABS operations.
        /// </summary>
        public object EtabsObject { get; set; }

        public object SapModel { get; set; }
    }
}
