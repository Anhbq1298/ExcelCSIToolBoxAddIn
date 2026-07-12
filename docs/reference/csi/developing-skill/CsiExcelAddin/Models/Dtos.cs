namespace CsiExcelAddin.Models
{
    /// <summary>
    /// Lightweight transfer object for a frame section read from a CSI model.
    /// Contains only the data needed by the ViewModel â€” not a CSI API type.
    /// </summary>
    public class FrameSectionDto
    {
        /// <summary>Section name as defined in the CSI model.</summary>
        public string Name { get; set; }

        /// <summary>Material name assigned to this section.</summary>
        public string Material { get; set; }

        /// <summary>Section depth in model units.</summary>
        public double Depth { get; set; }

        /// <summary>Section width in model units.</summary>
        public double Width { get; set; }
    }

    /// <summary>
    /// Lightweight transfer object for a joint (point object) in a CSI model.
    /// Coordinates are in the model's current length units.
    /// </summary>
    public class JointDto
    {
        /// <summary>Joint name or label as defined in the CSI model.</summary>
        public string Name { get; set; }

        public double X { get; set; }
        public double Y { get; set; }
        public double Z { get; set; }
    }

    /// <summary>
    /// Carries a status message from a service operation back to the ViewModel.
    /// Used as a consistent return type to avoid throwing for expected outcomes.
    /// </summary>
    public class OperationResult
    {
        public bool Success { get; }
        public string Message { get; }

        private OperationResult(bool success, string message)
        {
            Success = success;
            Message = message;
        }

        public static OperationResult Ok(string message = "Operation completed successfully.")
            => new OperationResult(true, message);

        public static OperationResult Fail(string message)
            => new OperationResult(false, message);
    }
}

