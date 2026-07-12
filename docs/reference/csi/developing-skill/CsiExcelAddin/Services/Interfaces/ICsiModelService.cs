using System.Collections.Generic;
using CsiExcelAddin.Models;

namespace CsiExcelAddin.Services.Interfaces
{
    /// <summary>
    /// Defines model-level operations shared across CSI products.
    /// Implementations live in product-specific adapters.
    /// Only methods that are verifiably available in all supported products
    /// belong here â€” product-unique operations go in the adapter directly.
    /// </summary>
    public interface ICsiModelService
    {
        /// <summary>
        /// Returns the name of the currently open model file.
        /// Useful for display in the ViewModel header.
        /// </summary>
        string GetModelFileName();

        /// <summary>
        /// Returns the current unit system set in the open model.
        /// Must be verified against the actual CSI API before implementation.
        /// </summary>
        string GetCurrentUnits();

        /// <summary>
        /// Returns all named frame sections defined in the model.
        /// Returns an empty list â€” not null â€” when nothing is defined.
        /// </summary>
        IReadOnlyList<FrameSectionDto> GetFrameSections();

        /// <summary>
        /// Returns all joint coordinates defined in the model.
        /// Returns an empty list â€” not null â€” when nothing is defined.
        /// </summary>
        IReadOnlyList<JointDto> GetJoints();
    }
}

