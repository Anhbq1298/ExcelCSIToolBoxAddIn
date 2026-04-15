namespace ExcelCSIToolBoxAddIn.Infrastructure.Etabs
{
    /// <summary>
    /// Value object containing UI display information for the currently attached ETABS model.
    /// </summary>
    public class EtabsAttachedModelInfo
    {
        public string ModelPath { get; set; }

        public string ModelDisplayText { get; set; }

        public string CurrentModelUnitText { get; set; }
    }
}
