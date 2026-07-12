using System.Collections.Generic;

namespace ExcelCSIToolBox.Application.ToolCatalog.Registry.Modules.CSI.FrameObject
{
    public static class FrameObjectIntentHints
    {
        public static readonly IReadOnlyList<string> Add = new[] { "add frame", "create beam", "draw member" };
        public static readonly IReadOnlyList<string> AssignSection = new[] { "assign section", "set frame property" };
        public static readonly IReadOnlyList<string> AssignDistributedLoad = new[] { "assign UDL", "assign distributed load", "apply uniform load", "load top chord", "top chord UDL" };
        public static readonly IReadOnlyList<string> Select = new[] { "select frame", "select beam" };
        public static readonly IReadOnlyList<string> GetSelected = new[] { "selected frames", "selected beams" };
        public static readonly IReadOnlyList<string> GetEndpoints = new[] { "frame endpoints", "frame points" };
        public static readonly IReadOnlyList<string> Delete = new[] { "delete frame", "remove beam" };
    }
}
