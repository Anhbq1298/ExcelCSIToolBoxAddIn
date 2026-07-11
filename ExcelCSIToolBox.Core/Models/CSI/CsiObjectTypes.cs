namespace ExcelCSIToolBox.Core.Models.CSI
{
    public static class CsiObjectTypes
    {
        public const string Point = "Point";
        public const string Frame = "Frame";
        public const string Area = "Area";
        public const string Pier = "Pier";
        public const string Spandrel = "Spandrel";
        public const string Unknown = "Unknown";

        public static string Normalize(string objectType)
        {
            if (string.IsNullOrWhiteSpace(objectType))
            {
                return Unknown;
            }

            string value = objectType.Trim();
            if (string.Equals(value, Point, System.StringComparison.OrdinalIgnoreCase) ||
                string.Equals(value, "Joint", System.StringComparison.OrdinalIgnoreCase))
            {
                return Point;
            }

            if (string.Equals(value, Frame, System.StringComparison.OrdinalIgnoreCase) ||
                string.Equals(value, "Line", System.StringComparison.OrdinalIgnoreCase))
            {
                return Frame;
            }

            if (string.Equals(value, Area, System.StringComparison.OrdinalIgnoreCase) ||
                string.Equals(value, "Shell", System.StringComparison.OrdinalIgnoreCase))
            {
                return Area;
            }

            if (string.Equals(value, Pier, System.StringComparison.OrdinalIgnoreCase))
            {
                return Pier;
            }

            if (string.Equals(value, Spandrel, System.StringComparison.OrdinalIgnoreCase))
            {
                return Spandrel;
            }

            return Unknown;
        }
    }
}
