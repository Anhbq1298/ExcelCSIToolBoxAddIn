using System.Collections.Generic;
using ExcelCSIToolBox.Core.Common.Results;

namespace ExcelCSIToolBox.Application.ToolCatalog.Contracts
{
    /// <summary>
    /// Application-layer entry point for AI and MCP tools that need to execute toolbox use cases.
    /// </summary>
    public interface IToolCatalogService
    {
        /// <summary>
        /// Returns the selected frame names from the active CSI model.
        /// </summary>
        OperationResult<IReadOnlyList<string>> GetSelectedFrameNames();
    }
}
