using System.Collections.Generic;

namespace ExcelCSIToolBox.Application.ToolCatalog.Registry.Modules.CSI.ShellObject
{
    public static class ShellObjectIntentHints
    {
        public static readonly IReadOnlyList<string> Add = new[] { "add shell", "add area", "create slab" };
        public static readonly IReadOnlyList<string> GetSelected = new[] { "selected shells", "selected areas" };
    }
}
