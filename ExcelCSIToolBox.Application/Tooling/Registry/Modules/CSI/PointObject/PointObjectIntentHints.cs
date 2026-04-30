using System.Collections.Generic;

namespace ExcelCSIToolBox.Application.Tooling.Registry.Modules.CSI.PointObject
{
    public static class PointObjectIntentHints
    {
        public static readonly IReadOnlyList<string> AddCartesian = new[] { "add point", "create point", "joint at coordinates" };
        public static readonly IReadOnlyList<string> Select = new[] { "select point", "select joint" };
        public static readonly IReadOnlyList<string> GetSelected = new[] { "selected points", "selected joints" };
        public static readonly IReadOnlyList<string> GetCoordinates = new[] { "point coordinates", "joint coordinates" };
        public static readonly IReadOnlyList<string> Delete = new[] { "delete point", "remove joint" };
    }
}
