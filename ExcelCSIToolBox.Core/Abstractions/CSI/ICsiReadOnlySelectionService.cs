using System.Collections.Generic;
using ExcelCSIToolBox.Core.Common.Results;
using ExcelCSIToolBox.Data;

namespace ExcelCSIToolBox.Core.Abstractions.CSI
{
    /// <summary>
    /// Read-only service for retrieving currently selected objects and frame names from
    /// a running ETABS or SAP2000 instance.
    /// The implementation must never expose the raw SapModel outside Infrastructure.
    /// </summary>
    public interface ICsiReadOnlySelectionService
    {
        /// <summary>
        /// Return all currently selected objects (points, frames, shells).
        /// Does not modify the model.
        /// </summary>
        OperationResult<List<CsiSelectedObjectDto>> GetSelectedObjects();

        /// <summary>
        /// Return unique names of currently selected frame objects only.
        /// Does not modify the model.
        /// </summary>
        OperationResult<List<string>> GetSelectedFrameNames();
    }

    /// <summary>
    /// Data transfer object representing a selected object in the CSI model.
    /// </summary>
    public class CsiSelectedObjectDto
    {
        /// <summary>Object type: "Point", "Frame", "Shell", or numeric id as string.</summary>
        public string ObjectType { get; set; }

        /// <summary>Unique name of the selected object.</summary>
        public string UniqueName { get; set; }
    }
}
