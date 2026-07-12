namespace CsiExcelAddin.Services.Interfaces
{
    /// <summary>
    /// Top-level contract that each CSI product adapter must satisfy.
    /// A concrete adapter (e.g. EtabsAdapter) composes both
    /// ICsiApplicationService and ICsiModelService and exposes them
    /// through this single entry point.
    ///
    /// Adding support for a new CSI product means implementing this interface
    /// â€” no changes required in the Ribbon, Views, or shared ViewModels.
    /// </summary>
    public interface ICsiProductAdapter
    {
        /// <summary>Friendly product name for display (e.g. "ETABS 21", "SAP2000 25").</summary>
        string ProductName { get; }

        /// <summary>Application-level service: attach, detach, connection state.</summary>
        ICsiApplicationService ApplicationService { get; }

        /// <summary>Model-level service: read sections, joints, units, etc.</summary>
        ICsiModelService ModelService { get; }
    }
}

