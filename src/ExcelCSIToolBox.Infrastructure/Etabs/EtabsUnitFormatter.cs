using System;

namespace ExcelCSIToolBox.Infrastructure.Etabs
{
    /// <summary>
    /// Formats ETABS database unit enums into human-readable strings.
    /// </summary>
    internal static class EtabsUnitFormatter
    {
        public static string FormatDatabaseUnits(
            ETABSv1.eForce forceUnits,
            ETABSv1.eLength lengthUnits,
            ETABSv1.eTemperature temperatureUnits)
        {
            string forceUnit = FormatForce(forceUnits);
            string lengthUnit = FormatLength(lengthUnits);
            string temperatureUnit = FormatTemperature(temperatureUnits);

            return $"{forceUnit}-{lengthUnit}-{temperatureUnit}";
        }

        private static string FormatForce(ETABSv1.eForce units)
        {
            switch (GetEnumKeyName(units).ToUpperInvariant())
            {
                case "KN": return "kN";
                case "KIP": return "kip";
                case "LB": return "lb";
                case "N": return "N";
                case "KGF": return "kgf";
                case "TONF": return "tonf";
                default: return GetEnumKeyName(units);
            }
        }

        private static string FormatLength(ETABSv1.eLength units)
        {
            switch (GetEnumKeyName(units).ToUpperInvariant())
            {
                case "M": return "m";
                case "MM": return "mm";
                case "CM": return "cm";
                case "FT": return "ft";
                case "INCH": return "inch";
                case "MICRON": return "micron";
                default: return GetEnumKeyName(units);
            }
        }

        private static string FormatTemperature(ETABSv1.eTemperature units)
        {
            switch (GetEnumKeyName(units).ToUpperInvariant())
            {
                case "C": return "C";
                case "F": return "F";
                default: return GetEnumKeyName(units);
            }
        }

        private static string GetEnumKeyName<TEnum>(TEnum enumValue) where TEnum : struct
        {
            var enumType = typeof(TEnum);
            var enumName = Enum.GetName(enumType, enumValue);
            if (!string.IsNullOrWhiteSpace(enumName))
            {
                return enumName;
            }

            if (!enumType.IsEnum)
            {
                return "?";
            }

            return Convert.ToInt32(enumValue).ToString();
        }
    }
}

