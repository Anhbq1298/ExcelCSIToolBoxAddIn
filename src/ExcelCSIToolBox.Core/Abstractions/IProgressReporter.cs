namespace ExcelCSIToolBox.Core.Abstractions
{
    /// <summary>
    /// Reports progress for long-running operations without coupling Core or Infrastructure to a UI.
    /// </summary>
    public interface IProgressReporter
    {
        void Report(int percent, string message);

        void ReportComplete(string message);

        void ReportError(string message);
    }
}
