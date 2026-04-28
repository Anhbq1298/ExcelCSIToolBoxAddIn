namespace ExcelCSIToolBoxAddIn.Infrastructure.Sap2000
{
    public static class Sap2000UnitFormatter
    {
        public static string FormatPresentUnits(SAP2000v1.eUnits units)
        {
            return units.ToString().Replace("_", "-");
        }
    }
}
