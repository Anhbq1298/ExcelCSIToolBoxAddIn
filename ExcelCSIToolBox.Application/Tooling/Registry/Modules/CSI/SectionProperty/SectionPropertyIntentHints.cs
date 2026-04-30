using System.Collections.Generic;

namespace ExcelCSIToolBox.Application.Tooling.Registry.Modules.CSI.SectionProperty
{
    public static class SectionPropertyIntentHints
    {
        public static readonly IReadOnlyList<string> GetList = new[] { "list sections", "section properties" };
        public static readonly IReadOnlyList<string> AssignToFrame = new[] { "assign section to frame", "set frame section" };
    }
}
