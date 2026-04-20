namespace CsiExcelAddin.Services.Interfaces
{
    /// <summary>
    /// Defines how the add-in attaches to and detaches from a running CSI application.
    /// Each CSI product (ETABS, SAP2000, SAFE) provides its own implementation
    /// through a product adapter — this interface keeps the UI layer product-agnostic.
    /// </summary>
    public interface ICsiApplicationService
    {
        /// <summary>True when the add-in holds an active connection to a CSI instance.</summary>
        bool IsAttached { get; }

        /// <summary>
        /// Attaches to a running CSI application process.
        /// Returns a result message suitable for display in the ViewModel status bar.
        /// </summary>
        AttachResult Attach();

        /// <summary>
        /// Releases the current CSI application reference cleanly.
        /// Safe to call even when not attached.
        /// </summary>
        void Detach();
    }

    /// <summary>Carries the outcome of an attach attempt back to the ViewModel.</summary>
    public class AttachResult
    {
        public bool Success { get; }
        public string Message { get; }

        public AttachResult(bool success, string message)
        {
            Success = success;
            Message = message;
        }

        public static AttachResult Ok(string message) => new AttachResult(true, message);
        public static AttachResult Fail(string message) => new AttachResult(false, message);
    }
}
