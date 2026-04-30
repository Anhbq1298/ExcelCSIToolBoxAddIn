using System.Collections.Generic;

namespace ExcelCSIToolBox.Application.Tooling.Registry.Modules.CSI.Model
{
    public static class ModelIntentHints
    {
        public static readonly IReadOnlyList<string> Info = new[] { "model file", "model path", "model info" };
        public static readonly IReadOnlyList<string> Units = new[] { "present units", "current units" };
    }
}
