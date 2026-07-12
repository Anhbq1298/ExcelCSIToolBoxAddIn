using System.Collections.Generic;

namespace ExcelCSIToolBox.Application.ToolCatalog.Registry.Modules.CSI.Selection
{
    public static class SelectionIntentHints
    {
        public static readonly IReadOnlyList<string> Clear = new[] { "clear selection", "unselect all" };
        public static readonly IReadOnlyList<string> GetSelectedObjects = new[] { "current selection", "selected objects" };
    }
}
