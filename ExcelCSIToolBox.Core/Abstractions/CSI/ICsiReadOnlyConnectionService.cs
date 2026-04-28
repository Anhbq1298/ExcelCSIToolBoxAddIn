using System.Collections.Generic;
using ExcelCSIToolBox.Core.Common.Results;
using ExcelCSIToolBox.Data;
using ExcelCSIToolBox.Data.DTOs.CSI;

namespace ExcelCSIToolBox.Core.Abstractions.CSI
{
    /// <summary>
    /// Read-only service for retrieving connection/model info and present units from a
    /// currently running ETABS or SAP2000 instance.
    /// The implementation must never expose the raw SapModel object outside Infrastructure.
    /// </summary>
    public interface ICsiReadOnlyConnectionService
    {
        /// <summary>Name of the product ("ETABS" or "SAP2000").</summary>
        string ProductName { get; }

        /// <summary>
        /// Attach to a running ETABS instance and return model info.
        /// Does not modify the model.
        /// </summary>
        OperationResult<CSISapModelConnectionInfoDTO> AttachToRunningEtabs();

        /// <summary>
        /// Attach to a running SAP2000 instance and return model info.
        /// Does not modify the model.
        /// </summary>
        OperationResult<CSISapModelConnectionInfoDTO> AttachToRunningSap2000();

        /// <summary>
        /// Return info about whichever instance is currently attached.
        /// Does not modify the model.
        /// </summary>
        OperationResult<CSISapModelConnectionInfoDTO> GetCurrentModelInfo();

        /// <summary>
        /// Return the current units of the attached model.
        /// Does not modify the model.
        /// </summary>
        OperationResult<string> GetPresentUnits();
    }
}
