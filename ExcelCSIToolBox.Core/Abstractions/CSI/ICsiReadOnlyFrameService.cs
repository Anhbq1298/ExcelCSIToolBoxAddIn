using System.Collections.Generic;
using ExcelCSIToolBox.Core.Common.Results;

namespace ExcelCSIToolBox.Core.Abstractions.CSI
{
    /// <summary>
    /// Read-only service for querying frame section assignments from a running CSI model.
    /// The implementation must never expose the raw SapModel outside Infrastructure.
    /// </summary>
    public interface ICsiReadOnlyFrameService
    {
        /// <summary>
        /// Return section names assigned to currently selected frame objects.
        /// Does not modify the model.
        /// </summary>
        OperationResult<List<FrameSectionAssignmentDto>> GetSelectedFrameSections();
    }

    /// <summary>
    /// DTO holding a frame name and its assigned section property name.
    /// </summary>
    public class FrameSectionAssignmentDto
    {
        /// <summary>Unique frame object name.</summary>
        public string FrameName { get; set; }

        /// <summary>Assigned section property name.</summary>
        public string SectionName { get; set; }
    }
}
